# Settings bridge v1

This is the implementation contract between the canonical AHK settings service and the Avalonia presentation layer. Bridge major version and schema major version are both 1. Ordinary fields must be exported from `SettingsSchema`; C# must not duplicate them.

## Transport

Normalize installation directory with absolute path, trim trailing slash, lowercase, encode UTF-8, then SHA256 to lowercase hexadecimal. A hidden top-level window has title `EVE-X-Preview.SettingsBridge/1/<hash>/runtime` (preferred) or `/offline`. A startup runtime endpoint advertises `ready:false` until its configuration is loaded. Discover hidden windows with Win32 FindWindow/EnumWindows (not visible-window filtering).

Send `WM_COPYDATA` with `dwData=0x45585031`, UTF-8 payload including terminating zero, maximum 4 MiB. `wParam` is a hidden sender window in the client process. Request:

```json
{"version":1,"id":"32 lowercase hexadecimal characters","method":"hello","session":"","arguments":{}}
```

Server validates the sender identity/current user, version, size, method, ID and session. It atomically writes `%LOCALAPPDATA%\EVE-X-Preview\IPC\<hash>\<id>.json`. Return value 1 means a response is available, 0 means rejected transport. Client uses SendMessageTimeout with 3 seconds maximum, then reads and removes its response. This directory contains only transient RPC responses, never user configuration. Expire abandoned responses. Filesystem errors must be surfaced. Bound response/request sizes and metadata without truncating JSON. Cached responses make an Apply request with the same ID idempotent. Responses:

```json
{"version":1,"id":"...","ok":true,"data":{}}
{"version":1,"id":"...","ok":false,"error":{"code":"validation","message":"...","details":[]}}
```

Errors include `incompatible`, `not-ready`, `busy`, `stale-session`, `validation`, `conflict`, `save-failed`, `runtime-unavailable`, `runtime-started`, `unknown-method`, `invalid-request`. Messages are useful for users. No arbitrary file paths/commands are accepted through RPC.

## Metadata

`metadata` returns `{bridgeVersion,schemaVersion,pages,sections,fields,tables,events,commands,fonts,themes}`. `pages` is an array of labels. Sections have `{id,page,title,help,group,fields:[field IDs],tables:[table kinds],commands:[command IDs]}`. Fields have `{id,label,kind,min,max,unit,help,path,group,global,section,dependency,choices:[{label,value}],default}`. Tables have `{kind,section,columns:[{label,kind}]}` derived from `SettingsTables`. Events have `{id,label,category}`. Commands have `{id,label,section,runtimeOnly,destructive}`. Specialized event descriptors are returned by `event`; ordinary field additions need no C# changes.

## Draft snapshot

`{session,profile,profiles:[{name,active}],dirty,clientDirty,values:{fieldId:{value,inherited,enabled}},canUndo,runtime,draft,inheritance:{sectionId:bool},creationInheritance:{sectionId:bool}}`. Scalar boolean values may be JSON booleans or numeric 0/1, matching legacy configuration; presentation flags are booleans. `draft` is opaque canonical recovery state, including installation identity, original/draft settings and profile-origin/undo metadata; only AHK interprets or merges it. Always preserve this state in client memory until an Apply result and reconnect are confirmed. Keep invalid unfinished edits in the client separately; do not replace them with old accepted values.

## Methods

Unless specified, mutations return a complete snapshot. `session` is required after `begin`.

| Method | Arguments | Result / behavior |
| --- | --- | --- |
| hello | `{}` | `{bridgeVersion,schemaVersion,role,installation,ready,pid}` |
| metadata | `{}` | Contract above |
| begin | `{clientPid,recovery?}` | Acquire lease; busy must preserve/focus existing draft; optional recovery is opaque `draft` |
| snapshot | `{}` | Current draft |
| client-state | `{dirty:bool}` | Report unfinished client edits so leases/runtime guards protect them even when no canonical mutation has been accepted |
| select-profile | `{name}` | Select edit profile, without runtime activation |
| set | `{id,value}` | Parse/validate one scalar using AHK |
| profile | `{action,name?,copyCurrent?,inheritance?}` | action add / rename / delete / activate; model semantics |
| inheritance | `{section,enabled}` | Toggle section inheritance |
| creation-inheritance | `{flags}` | Shared new-profile inheritance template |
| rows | `{kind,group?,rows?}` | Without rows: `{columns,rows,groups:[name],groupShortcuts:[{key,label,value}],inherited}`; with rows: validate/save through table adapter and return snapshot |
| group | `{kind,action,name?,newName?,key?,value?}` | kind characters/apps; action add/rename/delete/shortcut; canonical validation |
| event | `{id,values?}` | Without values: `{id,label,values,fields:[field descriptors],inherited}`; with values: validated canonical event mutation |
| sound | `{action,path?,volume?,events?}` | list returns `{rows:[{path,volume,assigned}],fields?}`; add/remove/gain/assign mutate via AHK, preserve event assignments/undo semantics |
| undo | `{}` | Canonical undo |
| validate | `{}` | `{issues:[{id,profile,message,severity}],conflicts:[path],canApply}` |
| apply | `{}` | `{applied:true,restartRequired,snapshot?}`; response flushed before restart; conflict/save errors preserve draft |
| discard | `{}` | Reset to latest runtime state |
| command | `{name,arguments?}` | Only enumerated commands; runtime-only operations unavailable in offline mode; actions with drafts must respect guards |
| end | `{discard:false}` | Release clean lease; dirty lease requires explicit `discard:true`; `{closed:true}` |

The command IDs are exported by `SettingsContract.Commands`: `test-thumbnail`, `stop-test-thumbnail`, `reshuffle-preview`, `reset-thumbnail-sizes`, `capture-client-placement`, `detected-sources`, `running-windows`, `shortcut-conflicts`, `show-log`, `export-diagnostics`, `toggle-live-previews`, `preview-sound`, `stop-sound`, `check-updates`, `update-status`, `install-stable`, `install-prerelease`, `changelog`, `documentation`, `report-problem`. Commands return a snapshot after draft mutations, otherwise `{message}` or dedicated results: `update-status` returns `{status,busy,stable,prerelease,installed}`; `detected-sources` returns `{game,chat}`; `running-windows` returns `{rows:[{exe,title,character}]}`; `shortcut-conflicts` returns `{groups}`. `preview-sound` accepts `{path,event?}`, where an event ID selects its effective sound/volume. `running-windows` accepts `{applications:bool}`.

## Availability and activation

`EVE-X-Preview-Settings.exe --probe <file>` writes `{application:"EVE-X-Preview-Settings",bridgeVersion:1,schemaVersion:1}` then exits without opening UI or connecting to AHK. Actual launch uses `--installation <directory>`. AHK bounds probe/startup/handshake and preserves internal fallback. Modern editor single-instance is per installation/current user; additional launches activate the existing window and exit.

Direct launch starts `SettingsHost.ahk` (source) or `EVE-X-Preview-SettingsHost.exe` (portable distribution) with `--parent <editor PID>` only if no runtime endpoint is available. This host does not initialize normal runtime behavior. Canonical read/save serialization and runtime-start handoff prevent competing writers. `--parent` is numeric, checked periodically, and no helper survives the parent.

The internal editor, external lease and runtime action guards share coordination. Runtime Apply restarts AHK after acknowledging, while the external process keeps its view state and reconnects with bounded retries. A lost connection must never silently discard local pending edits or the canonical recovery snapshot.

An active external draft or modal is focused when internal settings are requested. Clean exchange waits for a confirmed `end` or process exit; a refused close retains the lease. If the external process has exited with a canonical draft, opening internal settings offers explicit recovery, discard, or cancel. Recovery transfers the same canonical model, including profile origins and undo history, into the workbench. The prompt explains that unfinished text held only by the exited process cannot be recovered.

Custom interface palettes are exported in `themes` from `SettingsPalettes`. Each contains stable `id`, label/description and RGB roles: background, sidebar, input, foreground, muted, line, accent, selection. Both editors use the same definitions; ordinary choices are included in the canonical global SettingsTheme field.
