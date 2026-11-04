; Paint literal search matches without selecting or changing editable values.
class SettingsSearchText {
    static Matches(text, query) {
        if Trim(query) = ""
            return []
        text := text ""
        selected := Map()
        for word in StrSplit(Trim(query), " ") {
            if word = ""
                continue
            start := 1
            while start := InStr(text, word, false, start) {
                loop StrLen(word)
                    selected[start + A_Index - 1] := true
                start++
            }
        }
        ranges := [], index := 1
        while index <= StrLen(text) {
            if !selected.Has(index) {
                index++
                continue
            }
            start := index
            while selected.Has(index)
                index++
            ranges.Push({start: start, end: index - 1})
        }
        return ranges
    }
    static Width(dc, text) {
        size := Buffer(8, 0)
        DllCall("GetTextExtentPoint32W", "Ptr", dc, "WStr", text, "Int", StrLen(text), "Ptr", size)
        return NumGet(size, 0, "Int")
    }
    static Lines(dc, text, width, wrap) {
        lines := [], offset := 1
        for paragraph in StrSplit(text, "`n") {
            remaining := paragraph, consumed := 0
            if remaining = ""
                lines.Push({text: "", start: offset})
            while remaining != "" {
                length := StrLen(remaining), fit := length
                if wrap {
                    size := Buffer(8, 0)
                    DllCall("GetTextExtentExPointW", "Ptr", dc, "WStr", remaining, "Int", length,
                        "Int", Max(1, width), "Int*", &fit, "Ptr", 0, "Ptr", size)
                }
                take := Max(1, fit), used := take
                if take < length {
                    prefix := SubStr(remaining, 1, take)
                    space := InStr(prefix, " ", false, -1)
                    if space > 1
                        take := space - 1, used := space
                }
                lines.Push({text: SubStr(remaining, 1, take), start: offset + consumed})
                remaining := SubStr(remaining, used + 1)
                consumed += used
                while SubStr(remaining, 1, 1) = " " {
                    remaining := SubStr(remaining, 2)
                    consumed++
                }
            }
            offset += StrLen(paragraph) + 1
        }
        return lines
    }
    static Draw(dc, text, rect, flags, query, dark) {
        text := text ""
        DllCall("DrawTextW", "Ptr", dc, "WStr", text, "Int", -1, "Ptr", rect, "UInt", flags)
        if query = ""
            return 0
        text := StrReplace(text, "`r")
        if !(flags & 0x800) {
            text := StrReplace(text, "&&", Chr(0xE000))
            text := StrReplace(StrReplace(text, "&"), Chr(0xE000), "&")
        }
        matches := this.Matches(text, query)
        if !matches.Length
            return 0
        saved := DllCall("SaveDC", "Ptr", dc, "Int")
        left := NumGet(rect, 0, "Int"), top := NumGet(rect, 4, "Int")
        right := NumGet(rect, 8, "Int"), bottom := NumGet(rect, 12, "Int")
        DllCall("IntersectClipRect", "Ptr", dc, "Int", left, "Int", top, "Int", right, "Int", bottom)
        size := Buffer(8, 0)
        DllCall("GetTextExtentPoint32W", "Ptr", dc, "WStr", "Mg", "Int", 2, "Ptr", size)
        height := Max(1, NumGet(size, 4, "Int"))
        if flags & 4 && flags & 0x20
            top += Max(0, (bottom - top - height) // 2)
        lines := this.Lines(dc, text, right - left, !!(flags & 0x10) && !(flags & 0x20))
        DllCall("SetBkColor", "Ptr", dc, "UInt", SettingsTheme.BGR(dark ? "6B571A" : "FFE08A"))
        DllCall("SetTextColor", "Ptr", dc, "UInt", SettingsTheme.BGR(dark ? "FFF4C2" : "202020"))
        for index, line in lines {
            y := top + (index - 1) * height
            x := left
            width := this.Width(dc, line.text)
            if flags & 1
                x += (right - left - width) // 2
            else if flags & 2
                x := right - width
            for match in matches {
                start := Max(match.start, line.start), end := Min(match.end, line.start + StrLen(line.text) - 1)
                if start > end
                    continue
                prefix := SubStr(line.text, 1, start - line.start)
                value := SubStr(text, start, end - start + 1)
                runX := x + this.Width(dc, prefix)
                highlight := Buffer(16, 0)
                NumPut("Int", runX, "Int", y, "Int", runX + this.Width(dc, value), "Int", y + height, highlight)
                DllCall("ExtTextOutW", "Ptr", dc, "Int", runX, "Int", y, "UInt", 2,
                    "Ptr", highlight, "WStr", value, "UInt", StrLen(value), "Ptr", 0)
            }
        }
        DllCall("RestoreDC", "Ptr", dc, "Int", saved)
        return matches.Length
    }
}
