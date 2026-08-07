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

class TestRunnerClass {
    tests := []
    passed := 0
    failed := 0
    total := 0
    logPath := ""

    __New() {
        this.tests := []
        this.logPath := A_ScriptDir "\test_results.log"
        if FileExist(this.logPath)
            FileDelete(this.logPath)
    }

    Register(name, callback) {
        this.tests.Push(Map("name", name, "callback", callback))
    }

    Log(message) {
        FileAppend(message "`n", this.logPath)
    }

    RunAll() {
        this.passed := 0
        this.failed := 0
        this.total := this.tests.Length

        for item in this.tests {
            try {
                item["callback"].Call()
                this.passed++
                this.Log("PASS: " item["name"])
            } catch as err {
                this.failed++
                this.Log("FAIL: " item["name"])
                this.Log("  Message: " err.Message)
                if (err.File)
                    this.Log("  File: " err.File)
                if (err.Line)
                    this.Log("  Line: " err.Line)
            }
        }

        summary := "Test summary: " this.passed " passed, " this.failed " failed, " this.total " total."
        this.Log("")
        this.Log(summary)

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
#Include "Json_Functions_Test.ahk"

TestRunner.RunAll()
