; One modeless window shared by About, the tray and the post-update notice.
class ChangelogWindow {
    static window := 0

    static Show(afterUpdate := false, owner := 0) {
        if !IsObject(this.window) {
            window := Gui("+Resize +MinSize520x320", "EVE-X-Preview - Changelog")
            window.BackColor := "F7F7F7"
            window.SetFont("s12 w700", "Segoe UI")
            window.AddText("x16 y16 w728 h48 vHeading", "EVE-X-Preview v" UpdateChecker.CurrentVersion() "`nFull changelog")
            window.SetFont("s10 w400")
            contents := SimpleMarkdown.Add(window, "x16 y76 w728 h428 vContents")
            try SimpleMarkdown.SetText(contents, this.ReadText())
            catch as err {
                ProgramLog.Error(err, "Reading bundled changelog")
                SimpleMarkdown.SetText(contents, "The changelog could not be loaded. Please see the Changelog.md file in the project on GitHub.")
            }
            window.AddButton("x644 y520 w100 h28 Default vCloseButton", "Close").OnEvent("Click", (*) => this.Close())
            window.OnEvent("Close", (*) => this.Close())
            window.OnEvent("Escape", (*) => this.Close())
            window.OnEvent("Size", ObjBindMethod(this, "Resize"))
            this.window := window
            if owner {
                window.Opt("+Owner" owner)
                SettingsWindowSizing.ShowCentered(window, owner, 760, 564)
            } else
                window.Show("w760 h564")
            window["CloseButton"].Focus()
        } else {
            ; Repeated actions focus the same window, including when minimized.
            WinRestore(this.window.Hwnd)
            if owner {
                this.window.Opt("+Owner" owner)
                SettingsWindowSizing.CenterWindow(this.window.Hwnd, owner)
            }
            this.window.Show()
        }
        if afterUpdate
            this.window["Heading"].Value := "Successfully updated to EVE-X-Preview v" UpdateChecker.CurrentVersion() "`nWhat's new - full changelog"
        ; Rich Edit must be visible and laid out before its scroll range is final.
        SimpleMarkdown.ScrollToStart(this.window["Contents"])
        return this.window
    }

    static ReadText() {
        if !A_IsCompiled {
            SplitPath(A_LineFile, , &sourceDirectory)
            return FileRead(sourceDirectory "\..\Changelog.md", "UTF-8")
        }
        ; Read the release's own changelog directly from its executable, offline.
        module := DllCall("GetModuleHandleW", "Ptr", 0, "Ptr")
        resource := DllCall("FindResourceW", "Ptr", module, "Str", "CHANGELOG", "Ptr", 10, "Ptr") ; RT_RCDATA
        if !resource
            throw OSError()
        size := DllCall("SizeofResource", "Ptr", module, "Ptr", resource, "UInt")
        loaded := DllCall("LoadResource", "Ptr", module, "Ptr", resource, "Ptr")
        data := DllCall("LockResource", "Ptr", loaded, "Ptr")
        if !size || !data
            throw OSError()
        return LTrim(StrGet(data, size, "UTF-8"), Chr(0xFEFF))
    }

    static Resize(window, minMax, width, height) {
        if minMax = -1
            return
        window["Heading"].Move(, , width - 32)
        window["Contents"].Move(, , width - 32, height - 136)
        window["CloseButton"].Move(width - 116, height - 44)
    }

    static Close() {
        if IsObject(this.window)
            this.window.Destroy()
        this.window := 0
    }
}
