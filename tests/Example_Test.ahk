; Example test file. Add more test files in this folder and include them from TestRunner.ahk.

TestRunner.Register("Example arithmetic", ExampleArithmeticTest)

ExampleArithmeticTest() {
    AssertEqual(4, 2 + 2, "Basic addition should work.")
    AssertTrue(5 > 3, "Basic boolean checks should pass.")
    AssertFalse(1 = 2, "Basic inequality check should pass.")
}

TestRunner.Register("Example object key check", ExampleObjectKeyCheckTest)

ExampleObjectKeyCheckTest() {
    sample := Map("name", "EVE-X-Preview", "version", 1)
    AssertHasKey(sample, "name", "Sample object should include the name field.")
    AssertEqual("EVE-X-Preview", sample["name"], "Object values should match.")
}
