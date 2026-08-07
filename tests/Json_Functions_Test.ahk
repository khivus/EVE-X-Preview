; Tests for the JSON merge and migration helpers in Main.ahk.
; These are pure logic tests and do not require the GUI or EVE windows.

TestRunner.Register("JSON merge keeps user values", JsonMergeKeepsUserValuesTest)

JsonMergeKeepsUserValuesTest() {
    defaultObj := Map("a", 1, "nested", Map("x", 10, "y", 20), "list", [1, 2, 3])
    userObj := Map("a", 99, "nested", Map("x", 111), "list", [9, 9])

    result := JsonMergeNoOverwrite(defaultObj, userObj)

    AssertEqual(99, result["a"], "User value should win for primitive field.")
    AssertEqual(111, result["nested"]["x"], "User nested value should win.")
    AssertEqual(20, result["nested"]["y"], "Default nested field should be preserved.")
    AssertArrayEqual([9, 9], result["list"], "Arrays should remain user-owned and not be merged.")
}

TestRunner.Register("JSON merge adds missing default keys", JsonMergeAddsMissingDefaultKeysTest)

JsonMergeAddsMissingDefaultKeysTest() {
    defaultObj := Map("a", 1, "nested", Map("x", 10, "y", 20), "profile", Map("Default", Map("inner", "ok")))
    userObj := Map("nested", Map("x", 50))

    result := JsonMergeNoOverwrite(defaultObj, userObj)

    AssertEqual(1, result["a"], "Missing default primitive should be injected.")
    AssertEqual(50, result["nested"]["x"], "User value should stay.")
    AssertEqual(20, result["nested"]["y"], "Missing default nested value should be injected.")
}

TestRunner.Register("JSON merge adds missing profile entries", JsonMergeAddsMissingProfileEntriesTest)

JsonMergeAddsMissingProfileEntriesTest() {
    defaultObj := Map("_Profiles", Map(
        "Default", Map("name", "Default"),
        "Alpha", Map("name", "Alpha")
    ))
    userObj := Map("_Profiles", Map("Default", Map("name", "User Default")))

    result := JsonMergeNoOverwrite(defaultObj, userObj)

    AssertHasKey(result["_Profiles"], "Default", "Default profile should exist.")
    AssertHasKey(result["_Profiles"], "Alpha", "Default-only profile should be added.")
    AssertEqual("User Default", result["_Profiles"]["Default"]["name"], "User profile override should remain.")
    AssertEqual("Alpha", result["_Profiles"]["Alpha"]["name"], "Default profile should be copied.")
}

TestRunner.Register("DeepClone copies nested objects and arrays", DeepCloneCopiesNestedObjectsAndArraysTest)

DeepCloneCopiesNestedObjectsAndArraysTest() {
    source := Map("a", 1, "nested", Map("x", [1, 2, Map("z", 3)]))
    clone := DeepClone(source)

    clone["a"] := 99
    clone["nested"]["x"][3]["z"] := 123

    AssertEqual(1, source["a"], "Top-level primitive should not be mutated.")
    AssertEqual(3, source["nested"]["x"][3]["z"], "Nested object should not be shared.")
}

TestRunner.Register("IsArrayLike detects arrays", IsArrayLikeDetectsArraysTest)

IsArrayLikeDetectsArraysTest() {
    AssertTrue(IsArrayLike([1, 2, 3]), "Array literal should be recognized as array-like.")
    AssertFalse(IsArrayLike(Map("a", 1, "b", 2)), "Plain object should not be treated as array-like.")
    AssertFalse(IsArrayLike(Map("Length", 7)), "Map-like objects must not be mistaken for arrays just because they have a Length field.")
}

TestRunner.Register("JsonMergeNoOverwrite handles object-vs-primitive mismatches", JsonMergeNoOverwriteHandlesObjectVsPrimitiveMismatchesTest)

JsonMergeNoOverwriteHandlesObjectVsPrimitiveMismatchesTest() {
    defaultObj := Map("nested", Map("x", 1), "plain", 7)
    userObj := Map("nested", 99, "plain", Map("still", "user"))

    result := JsonMergeNoOverwrite(defaultObj, userObj)

    AssertDeepEqual(Map("x", 1), result["nested"], "Default object should win when user has a primitive at the same key.")
    AssertDeepEqual(Map("still", "user"), result["plain"], "User object should be preserved when default is primitive.")
}

TestRunner.Register("JsonMergeNoOverwrite keeps nested arrays independent", JsonMergeNoOverwriteKeepsNestedArraysIndependentTest)

JsonMergeNoOverwriteKeepsNestedArraysIndependentTest() {
    defaultObj := Map("items", [Map("id", 1), Map("id", 2)], "outer", Map("inner", [1, 2, 3]))
    userObj := Map("items", [Map("id", 99)], "outer", Map("inner", [9, 9]))

    result := JsonMergeNoOverwrite(defaultObj, userObj)

    AssertArrayEqual([Map("id", 99)], result["items"], "User-owned arrays should not be merged with defaults.")
    AssertArrayEqual([9, 9], result["outer"]["inner"], "Nested user arrays should remain intact.")
}

TestRunner.Register("MigrateSettings migrates legacy global settings", MigrateSettingsMigratesLegacyGlobalSettingsTest)

MigrateSettingsMigratesLegacyGlobalSettingsTest() {
    userObj := Map(
        "global_Settings", Map(
            "LastUsedProfile", "Main",
            "First_Start_After_Update", true,
            "ThisThat", "legacy"
        ),
        "_Profiles", Map("Main", Map("name", "Main"))
    )

    migrated := MigrateSettings(userObj, "", 1)

    AssertEqual("Main", migrated["LastUsedProfile"], "LastUsedProfile should migrate.")
    AssertEqual(true, migrated["First_Start_After_Update"], "First_Start_After_Update should migrate.")
    AssertEqual("legacy", migrated["ThisThat"], "ThisThat should migrate.")
    AssertHasKey(migrated["_Profiles"], "Main", "Profiles should remain available after migration.")
}

TestRunner.Register("MigrateSettings fails on invalid legacy structure", MigrateSettingsFailsOnInvalidLegacyStructureTest)

MigrateSettingsFailsOnInvalidLegacyStructureTest() {
    userObj := Map("global_Settings", Map("LastUsedProfile", "Main"))
    AssertThrows(TestCallback, "No profiles found", "Missing profile structure should throw.")
}

TestCallback() {
    MigrateSettings(Map("global_Settings", Map("LastUsedProfile", "Main")), "", 1)
}
