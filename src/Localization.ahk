; English UI strings are the fallback and the catalog keys. Keep stored values,
; command IDs, and user-supplied text outside these lookups.
T(text, args*) {
    translated := Localization.Messages.Get(text, text)
    if !args.Length
        return translated
    try return Format(translated, args*)
    catch
        return Format(text, args*)
}

class Localization {
    static Messages := Map()

    static Initialize(language, installation := A_ScriptDir) {
        catalogs := Map()
        try {
            if A_IsCompiled {
                module := DllCall("GetModuleHandleW", "Ptr", 0, "Ptr")
                resource := DllCall("FindResourceW", "Ptr", module, "Str", "LOCALES", "Ptr", 10, "Ptr")
                if resource {
                    size := DllCall("SizeofResource", "Ptr", module, "Ptr", resource, "UInt")
                    loaded := DllCall("LoadResource", "Ptr", module, "Ptr", resource, "Ptr")
                    data := DllCall("LockResource", "Ptr", loaded, "Ptr")
                    if size && data
                        catalogs := JSON.Load(LTrim(StrGet(data, size, "UTF-8"), Chr(0xFEFF)))
                }
            } else {
                SplitPath(A_LineFile, , &sourceDirectory)
                catalogs := JSON.Load(FileRead(sourceDirectory "\..\res\locales.json", "UTF-8"))
            }
        }
        try {
            overrides := JSON.Load(FileRead(installation "\EVE-X-Preview.locales.json", "UTF-8"))
            for id, entry in overrides {
                if !(entry is Map) || !RegExMatch(id, "^[a-z]{2,3}(-[a-zA-Z0-9]{2,8})*$")
                    continue
                if !catalogs.Has(id)
                    catalogs[id] := Map("name", id, "messages", Map())
                if entry.Get("name", "") is String && entry.Get("name", "") != ""
                    catalogs[id]["name"] := entry["name"]
                if entry.Get("messages", 0) is Map
                    for key, value in entry["messages"]
                        if value is String && value != ""
                            catalogs[id]["messages"][key] := value
            }
        }
        choices := [{label: "English", value: "en"}]
        for id, entry in catalogs
            if id != "en"
                choices.Push({label: entry["name"], value: id})
        SettingsSchema.Fields["InterfaceLanguage"].choices := choices
        this.Messages := language is String && catalogs.Has(language) ? catalogs[language]["messages"] : Map()
    }
}

class LocalizedSettingsGui extends Gui {
    __New(options := "", title := "") => super.__New(options, T(title))
    AddText(options := "", text := "") => super.AddText(options, T(text))
    AddButton(options := "", text := "") => super.AddButton(options, T(text))
    AddCheckbox(options := "", text := "") => super.AddCheckbox(options, T(text))
    AddGroupBox(options := "", text := "") => super.AddGroupBox(options, T(text))
    AddRadio(options := "", text := "") => super.AddRadio(options, T(text))
}
