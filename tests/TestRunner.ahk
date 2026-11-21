#Requires AutoHotkey v2.0
#SingleInstance Force
; SetBatchLines(-1)
SetWinDelay -1
SetWorkingDir(A_ScriptDir)

; This runner is the shared test harness for all test files in this folder.
; Add new test files and include them here, then run this file.

global __EVE_X_PREVIEW_TESTING__ := true
#Include "..\Main.ahk"

global TestRunner := TestRunnerClass()
OnExit((*) => TestRunner.CloseLog())
SettingsBackup.Root := A_ScriptDir "\..\.vs\test-settings-backups"
if A_Args.Length && A_Args[1] = "--headless"
    OnError(HeadlessTestError)
HeadlessTestError(err, mode) {
    FileAppend("Unhandled test error: " err.Message " at " err.File ":" err.Line "`n", "**")
    ExitApp(2)
}

class TestRunnerClass {
    tests := []
    passed := 0
    failed := 0
    total := 0
    logPath := ""

    __New(logPath := "") {
        this.tests := []
        this.logPath := logPath != "" ? logPath : A_ScriptDir "\test_results.log"
        ; Keep write access for the entire run. Snapshot readers must share it;
        ; they cannot acquire a deny-write handle between individual log lines.
        this.logFile := FileOpen(this.logPath, "w", "UTF-8")
    }

    Register(name, callback) {
        this.tests.Push(Map("name", name, "callback", callback))
    }

    Log(message) {
        if !IsObject(this.logFile)
            throw Error("The test log is closed.")
        text := message "`n"
        if this.logFile.Write(text) != StrPut(text, "UTF-8") - 1
            throw Error("Incomplete write to the test log.")
        ; Accessing Handle flushes AHK's buffer so live readers see each line.
        if !DllCall("FlushFileBuffers", "Ptr", this.logFile.Handle)
            throw OSError()
    }

    CloseLog() {
        if IsObject(this.logFile) {
            this.logFile.Close()
            this.logFile := 0
        }
    }

    RunAll() {
        this.passed := 0
        this.failed := 0
        selected := this.tests
        if A_Args.Length >= 3 && A_Args[2] = "--filter" {
            selected := []
            for item in this.tests
                if InStr(item["name"], A_Args[3])
                    selected.Push(item)
            if !selected.Length {
                FileAppend("No tests match the requested filter.`n", "*")
                ExitApp(2)
            }
        }
        this.total := selected.Length

        for item in selected {
            this.Log("RUN: " item["name"])
            failure := 0
            try {
                item["callback"].Call()
            } catch as err {
                failure := err
            }
            ; Logging failures are runner errors, not failed test assertions.
            if IsObject(failure) {
                this.failed++
                this.Log("FAIL: " item["name"])
                this.Log("  Message: " failure.Message)
                if (failure.File)
                    this.Log("  File: " failure.File)
                if (failure.Line)
                    this.Log("  Line: " failure.Line)
            } else {
                this.passed++
                this.Log("PASS: " item["name"])
            }
        }

        summary := "Test summary: " this.passed " passed, " this.failed " failed, " this.total " total."
        this.Log("")
        this.Log(summary)
        this.CloseLog()

        if A_Args.Length && A_Args[1] = "--headless" {
            FileAppend(summary "`n", "*")
            ExitApp(this.failed > 0 ? 1 : 0)
        }

        if (this.failed > 0) {
            MsgBox(summary "`nSee log: " this.logPath, "Test Runner", "Icon! 4096")
            ExitApp(1)
        }

        MsgBox(summary "`nLog: " this.logPath, "Test Runner", "Iconi 4096")
        ExitApp(0)
    }
}

AssertTrue(condition, message := "Expected condition to be true.") {
    if (!condition)
        throw Error(message)
}

AssertFalse(condition, message := "Expected condition to be false.") {
    if (condition)
        throw Error(message)
}

AssertArrayEqual(expected, actual, message := "Arrays are not equal.") {
    if (!IsArrayLike(expected) || !IsArrayLike(actual) || expected.Length != actual.Length) {
        throw Error(message " (expected: " expected ", actual: " actual ")")
    }

    for index, value in expected {
        if (IsObject(value) || IsObject(actual[index])) {
            AssertDeepEqual(value, actual[index], message)
            continue
        }

        if (value != actual[index])
            throw Error(message " (expected: " expected ", actual: " actual ")")
    }
}

AssertDeepEqual(expected, actual, message := "Values are not equal.") {
    if (IsObject(expected) || IsObject(actual)) {
        if (!IsObject(expected) || !IsObject(actual))
            throw Error(message " (expected: " expected ", actual: " actual ")")

        if (Type(expected) != Type(actual))
            throw Error(message " (expected: " expected ", actual: " actual ")")

        if (expected is Array) {
            if (expected.Length != actual.Length)
                throw Error(message " (expected: " expected ", actual: " actual ")")

            for index, value in expected {
                AssertDeepEqual(value, actual[index], message)
            }
            return
        }

        if (expected is Map) {
            if (expected.Count != actual.Count)
                throw Error(message " (expected: " expected ", actual: " actual ")")

            for key, value in expected {
                if (!actual.Has(key))
                    throw Error(message " (expected: " expected ", actual: " actual ")")
                AssertDeepEqual(value, actual[key], message)
            }
            return
        }

        if (expected != actual)
            throw Error(message " (expected: " expected ", actual: " actual ")")
        return
    }

    if (expected != actual)
        throw Error(message " (expected: " expected ", actual: " actual ")")
}

AssertEqual(expected, actual, message := "Values are not equal.") {
    AssertDeepEqual(expected, actual, message)
}

AssertNotEqual(expected, actual, message := "Values should not be equal.") {
    if (expected = actual)
        throw Error(message)
}

AssertHasKey(obj, key, message := "Expected key was not found.") {
    if (!IsObject(obj) || !obj.Has(key))
        throw Error(message " (key: " key ")")
}

AssertThrows(callback, expectedText := "", message := "Expected callback to throw.") {
    try {
        callback.Call()
    } catch as err {
        if (expectedText != "" && !InStr(err.Message, expectedText))
            throw Error(message " (expected to contain: " expectedText ", got: " err.Message ")")
        return
    }

    throw Error(message)
}

#Include "Example_Test.ahk"
#Include "Runner_Paint_Regression_Test.ahk"
#Include "Json_Functions_Test.ahk"
#Include "Thumbnail_Names_Test.ahk"
#Include "Thumbnail_Event_Text_Test.ahk"
#Include "Combat_Events_Test.ahk"
#Include "NPC_Database_Test.ahk"
#Include "System_Tracking_Test.ahk"
#Include "System_Colors_Test.ahk"
#Include "DPS_Monitoring_Test.ahk"
#Include "Game_Log_Discovery_Test.ahk"
#Include "Thumbnail_Drag_Test.ahk"
#Include "Log_Worker_Test.ahk"
#Include "Update_Checker_Test.ahk"
#Include "Changelog_Test.ahk"
#Include "Group_Activation_Test.ahk"
#Include "Preview_Diagnostics_Test.ahk"
#Include "Program_Log_Memory_Test.ahk"
#Include "Thumbnail_Shift_Test.ahk"
#Include "Profile_Switch_Test.ahk"
#Include "Hotkey_Settings_Test.ahk"
#Include "Settings_Layout_Test.ahk"
#Include "Color_Settings_Test.ahk"
#Include "Settings_Thumbnail_Preview_Test.ahk"
#Include "Thumbnail_Border_Test.ahk"
#Include "Thumbnail_Capture_Test.ahk"
#Include "Event_Sounds_Test.ahk"

#Include "Thumbnail_Tools_Test.ahk"
#Include "Settings_Scroll_Test.ahk"
#Include "Settings_Redesign_Model_Test.ahk"
#Include "Settings_Workbench_Test.ahk"
#Include "Settings_Grid_Test.ahk"
#Include "Settings_Interaction_Followup_Test.ahk"
#Include "Settings_Appearance_Test.ahk"
#Include "Settings_Polish_Test.ahk"
#Include "Settings_Chrome_Test.ahk"
#Include "Settings_Paint_Lifecycle_Test.ahk"
#Include "Settings_Scroll_Rendering_Test.ahk"
#Include "Settings_Input_Pulse_Test.ahk"
#Include "Settings_Review_Followup_Test.ahk"
#Include "Settings_Release_Safety_Test.ahk"
#Include "Settings_Followup_Test.ahk"
#Include "Thumbnail_Recovery_Test.ahk"
#Include "Release_Packaging_Test.ahk"
#Include "Settings_Contract_Test.ahk"
#Include "Settings_Bridge_Test.ahk"
#Include "Update_Handoff_Test.ahk"
#Include "Settings_Editor_Test.ahk"
#Include "Settings_Themes_Test.ahk"
#Include "Update_1704_Test.ahk"
#Include "Character_List_Test.ahk"
TestRunner.RunAll()
