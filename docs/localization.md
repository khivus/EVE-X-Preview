# Interface translations

Choose **Application → Settings appearance → Interface language** in either Settings editor. Apply the change and reopen Settings. In Russian, the path is **Приложение → Вид и язык настроек → Язык интерфейса**. The language applies across profiles and to the tray menu.

## Add or edit a translation

Save a UTF-8 file named `EVE-X-Preview.locales.json` next to the app and restart it. Each top-level key is a language tag such as `de` or `pt-BR`. The name appears in the language picker. Use the built-in [Russian catalog](../res/locales.json) to find English message keys.

```json
{
  "de": {
    "name": "Deutsch",
    "messages": {
      "Open": "Einstellungen öffnen",
      "Statistics": "Statistik",
      "Profile activated: {1}": "Profil aktiviert: {1}"
    }
  }
}
```

To update Russian wording, use the tag `ru` and list the messages you want to override. You can omit the name for an existing language. Missing overrides use the bundled translation when available, then the English text. To restore bundled wording, remove the override and restart.

Copy keys exactly, including capitalization, punctuation, and whitespace. Preserve placeholders such as `{0}` and `{1}`. Keep tray menu labels distinct so Windows can identify each menu item. The catalog affects UI messages. Profile and character names, stored values, and game log inputs retain their original text.

## Maintain the bundled catalog

Edit [`res/locales.json`](../res/locales.json) and rebuild the package. The app and offline Settings host embed the catalog, and modern Settings gets the active messages from its host. For new interface text, use `T()` at the point where it is displayed. Built-in Settings controls and the common modern Settings text helpers already apply the lookup. Keep settings keys, bridge identifiers, choice values, and user input independent of the displayed wording.

After updating a language, test its picker, both Settings editors, translated search, and the tray menu. Check longer labels at a narrow window size. Try a message containing a profile name or a number to confirm its placeholders still render.
