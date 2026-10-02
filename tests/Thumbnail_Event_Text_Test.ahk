TestRunner.Register("Event text is independent of the custom thumbnail name", ThumbnailEventTextTest)
ThumbnailEventTextTest() {
    app := ThumbnailNamesFixture()
    app.CustomThumbnailNames := Map("Characters", "Alice", "Names", "Scout")
    nameText := {Text: "Scout"}
    eventText := {Text: ""}
    window := {Title: "EVE - Alice"}
    app.ThumbWindows.test := Map("TextOverlay", Map("OverlayText", nameText, "EventText", eventText, "SystemText", {Text: ""}), "Window", window)
    app.updateThumbnailEventText("Under Attack By Player", "test")
    AssertEqual("Under Attack By Player", eventText.Text)
    AssertEqual("Scout", nameText.Text)
    app.SetThumbnailText["test"] := "EVE - Alice"
    AssertEqual("Under Attack By Player", eventText.Text, "Refreshing the name must preserve an active event.")
    app.updateThumbnailEventText("Warp Disrupted", "test")
    AssertEqual("Warp Disrupted", eventText.Text)
    app.updateThumbnailEventText("", "test")
    AssertEqual("", eventText.Text)
    AssertEqual("Scout", nameText.Text)
    app.updateThumbnailEventText("Warp Disrupted", "test")
    app.SetThumbnailText["test"] := "EVE - Bob"
    AssertEqual("", eventText.Text, "A different character must not inherit the old event.")
    AssertEqual("Bob", nameText.Text)
    app.ThumbWindows.DeleteProp("test")
    app.updateThumbnailEventText("", "test") ; Expired timer after the thumbnail closed.
}

TestRunner.Register("Event utility line stays inside resized thumbnails", ThumbnailEventLayoutTest)
ThumbnailEventLayoutTest() {
    app := ThumbnailNamesFixture()
    overlay := Gui("-Caption")
    try {
        overlay.MarginX := 5
        overlay.MarginY := 5
        overlay.SetFont("s12", "Gill Sans MT")
        app.AddThumbnailTextControls(overlay, "EVE - Alice", 200, 150)
        for size in [[200, 150], [400, 300], [50, 50], [8, 8]] {
            app.LayoutThumbnailText(overlay, 0, size[1], size[2])
            overlay["OverlayText"].GetPos(&nx, &ny, &nw, &nh)
            overlay["EventText"].GetPos(&ex, &ey, &ew, &eh)
            AssertTrue(ex >= 0 && ey >= 0 && ex + ew <= size[1] && ey + eh <= size[2], "Event text must stay inside the thumbnail.")
            AssertTrue(ny + nh <= ey, "Name and event areas must not overlap.")
            AssertEqual(size[2] - Min(5, size[2] / 2), ey + eh, "Event line must follow the bottom margin.")
        }
    } finally {
        overlay.Destroy()
    }
}

TestRunner.Register("Wrapped character names take priority over event messages", ThumbnailNamePriorityTest)
ThumbnailNamePriorityTest() {
    for trackSystem in [0, 1] {
        app := ThumbnailNamesFixture()
        app.systemTrackingEnabled := trackSystem
        overlay := Gui("-Caption")
        try {
            overlay.MarginX := 5
            overlay.MarginY := 5
            overlay.SetFont("s12", "Gill Sans MT")
            app.AddThumbnailTextControls(overlay, "EVE - Rakkadan Alihaken", 100, 150)
            app.ThumbWindows.test := Map("TextOverlay", overlay)
            overlay["EventText"].Text := "Fleet Warped"
            overlay["SystemText"].Text := "Dodixie"
            nameHeight := app.MeasureThumbnailNameHeight(overlay, 90)
            AssertTrue(nameHeight > overlay.EventTextHeight, "The sample name must wrap to multiple lines.")
            smallHeight := nameHeight + 10 + (trackSystem ? overlay.EventTextHeight : 0)
            app.LayoutThumbnailText(overlay, 0, 100, smallHeight)
            overlay["OverlayText"].GetPos(, &ny, , &nh)
            overlay["EventText"].GetPos(, &ey, , &eh)
            AssertTrue(nh >= nameHeight, "The full wrapped name must retain its height.")
            AssertEqual(0, eh, "Hide the event when it would occupy the name's space.")
            AssertTrue(ey >= ny + nh)
            app.LayoutThumbnailText(overlay, 0, 100, smallHeight + overlay.EventTextHeight)
            overlay["EventText"].GetPos(, , , &eh)
            AssertEqual(overlay.EventTextHeight, eh, "Enlarging must restore the event line.")
            AssertEqual("Fleet Warped", overlay["EventText"].Text)
            app.LayoutThumbnailText(overlay, 0, 100, smallHeight)
            app.updateThumbnailText("EVE - Al", "test")
            overlay["EventText"].GetPos(, , , &eh)
            AssertTrue(eh > 0, "A shorter character name must release space for the event.")
        } finally {
            overlay.Destroy()
        }
    }
}
