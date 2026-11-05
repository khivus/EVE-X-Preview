; Canonical interface palettes shared by native AHK and exported to Avalonia.
; These colors never modify thumbnail, character, or monitored-event colors.
class SettingsPalettes {
    static Order := ["graphite", "matrix", "kitsune", "synthwave", "cyberpunk"]
    static Items := Map(
        "graphite", Map("id", "graphite", "label", "Graphite", "description", "Neutral charcoal and soft gray.",
            "background", "242628", "sidebar", "202224", "input", "303336", "foreground", "E4E5E7",
            "muted", "B0B3B8", "line", "484C50", "accent", "AEBECE", "selection", "444D57"),
        "matrix", Map("id", "matrix", "label", "The Matrix", "description", "Terminal greens and digital rain.",
            "background", "041008", "sidebar", "020B06", "input", "0A2013", "foreground", "AAF6B1",
            "muted", "93C79C", "line", "1B4528", "accent", "47C000", "selection", "082F14"),
        "kitsune", Map("id", "kitsune", "label", "Kitsune", "description", "Fox-spirit rose, wine and midnight charcoal.",
            "background", "0E0E12", "sidebar", "150C13", "input", "200E17", "foreground", "F8E9EE",
            "muted", "CCB4C0", "line", "4D2A3E", "accent", "FF7AAB", "selection", "64173B"),
        "synthwave", Map("id", "synthwave", "label", "Synthwave", "description", "Neon pink over retro violet and purple.",
            "background", "130C22", "sidebar", "170B25", "input", "2F1748", "foreground", "FAE6FF",
            "muted", "C9ABD8", "line", "533064", "accent", "FF4BD2", "selection", "3B1659"),
        "cyberpunk", Map("id", "cyberpunk", "label", "Cyberpunk", "description", "Electric magenta, midnight blue and teal.",
            "background", "080B18", "sidebar", "0D1021", "input", "181832", "foreground", "F4F1FF",
            "muted", "A9B8CD", "line", "303557", "accent", "FF49C4", "selection", "007052"))

    static Export() {
        palettes := []
        for id in this.Order
            palettes.Push(this.Items[id].Clone())
        return palettes
    }
}
