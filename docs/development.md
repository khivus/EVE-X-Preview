# Development

[Back to README](../README.MD)

## Run from source

1. Download and install [AutoHotkey v2](https://www.autohotkey.com/).
2. Download the ZIP archive from the [releases page](https://github.com/khivus/EVE-X-Preview/releases).
3. Extract the archive and run `Main.ahk`.

## Contributing

Contributions are welcome. All pull requests must target the [`dev` branch](https://github.com/khivus/EVE-X-Preview/tree/dev).

1. Fork the repository and use the latest upstream `dev` as the starting point whenever possible. If you already have a fork, sync its `dev` branch with this repository before starting.
2. Create a separate branch for your changes from the updated `dev` branch.
3. Keep changes focused, test the affected behavior, and describe what changed and how you tested it in the pull request.
4. Before submitting, update your branch with the latest upstream `dev` whenever possible and resolve any conflicts.
5. Open the pull request against `khivus/EVE-X-Preview` with **base: `dev`**.

### EVE-X-Preview

* Created by [gonzo83](https://github.com/g0nzo83/EVE-X-Preview)
* Forked and maintained by [khivus](https://github.com/khivus)
* Credits to [mrmjstc](https://github.com/mrmjstc/eve-maj-preview) and [CJKondur](https://github.com/CJKondur/EVE-MultiPreview)

## Source layout and tests

`Main.ahk` starts the app, `src/` holds application code, `Lib/` holds libraries and default settings, and `res/` holds icons used by the script and Ahk2Exe. `Changelog.md` is embedded in compiled releases for the offline changelog viewer.

Run the Windows AutoHotkey v2 suite from the repository root:

```powershell
& 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe' /ErrorStdOut tests/TestRunner.ahk --headless
```

Tests use temporary windows and fixture settings. Live EVE checks and rebuilding a release are separate verification steps.
