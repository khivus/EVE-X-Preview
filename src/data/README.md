# NPC name data

`NPCGeneralNames.ahk` is generated from the [EVE Ref bulk reference dataset](https://docs.everef.net/datasets/reference-data.html), which combines CCP's SDE, ESI and Hoboleaks data. It includes every English type name in **Entity category 11**, including structures, special NPCs and unpublished types. Publication is not a usable NPC filter: the 2026-09-10 snapshot has 7,378 unpublished entities out of 7,386 total.

The generated header records the source build timestamp, input SHA256 and counts. Identical names are deduplicated; the case-insensitive runtime Map also merges case variants. This snapshot has 6,855 distinct English names. It is a complete category snapshot, not a guarantee that every dynamically assigned in-game display name appears in static data. Test/obsolete entities are retained rather than guessed from their names. English names match the application's existing English combat-log parser.

`../NPCDatabase.ahk` owns the curated faction, officer and capital lists, copied without membership changes. It also owns explicit display-name aliases supported by combat-log evidence. General recognition includes these names when the corresponding special event is disabled. Special event priority remains faction, officer, capital, then general.

The AHK files are included in the application and compiled executable. Maps are created once on first use and shared across monitoring restarts. Combat checks use `Map.Has()` with no name-list scan, disk access or network request. Treat the shared maps as read-only. General matching ignores case; the special maps retain their previous case-sensitive behavior. Tagged player names do not match, except that complete reference names such as `[AIR] Incursus` remain valid NPCs. A player with exactly the same untagged name as an NPC cannot be distinguished by a name lookup alone.

## Refreshing

Requirements: Windows PowerShell 5.1 or later and 7-Zip. From the repository root:

```powershell
$npcDataDirectory = Join-Path $env:TEMP ('eve-npc-data-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $npcDataDirectory | Out-Null
$npcArchive = Join-Path $npcDataDirectory 'reference-data.tar.xz'
curl.exe --fail --location --output $npcArchive https://data.everef.net/reference-data/reference-data-latest.tar.xz
& 'C:\Program Files\7-Zip\7z.exe' e $npcArchive "-o$npcDataDirectory" -y
& 'C:\Program Files\7-Zip\7z.exe' e (Join-Path $npcDataDirectory 'reference-data.tar') meta.json groups.json types.json "-o$npcDataDirectory" -y
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Update-NPCNames.ps1 -ReferenceDirectory $npcDataDirectory
& 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe' /ErrorStdOut .\tests\TestRunner.ahk --headless
git diff --check
```

Check each command succeeds before continuing. The generator reads the large `types.json` file into memory, so allow several GB of free memory during regeneration; this cost does not apply to the application. Reusing the same extracted dataset produces identical output. Review and commit the generated AHK file; do not commit the downloaded archive or JSON inputs. No automatic runtime updates occur, and regeneration never changes the curated special lists.

EVE Online names and game data are the intellectual property of CCP hf. EVE Ref is maintained by Autonomous Logic. These data sources do not endorse this project.

## Incoming damage profiles

`DamageProfiles.ahk` is generated independently by `tools/Update-DamageProfiles.ps1`. It contains exact English names and EM/thermal/kinetic/explosive damage profiles only for drones (category 18), fighters (category 87), and smartbomb exceptions (group 72). The 2026-09-10 build produces 178 names with two ambiguous names omitted. Missiles use four shared damage-family mappings in `DPSMeter`: Mjolnir = EM, Inferno = thermal, Scourge = kinetic, Nova = explosive. Family names are matched as whole words in the weapon field, covering faction and advanced variants without individual item entries. Smartbombs likewise use EMP = EM, Plasma = thermal, Graviton = kinetic, Proton = explosive, and Multispectrum = equal parts of all four. The generator verifies these smartbomb mappings against the source damage attributes before omitting their rows. Micro YF-12a Smartbomb retains an exact thermal profile because its name lacks a damage family. Turret ammunition, crystals, bombs, and scripts are excluded. It does not change any NPC list.

The generator verifies dogma attributes 114, 118, 117, and 116 against their names (`emDamage`, `thermalDamage`, `kineticDamage`, and `explosiveDamage`). Fighters additionally use their turret-attack, missile-attack, missile-ability and kamikaze damage attributes (2171-2174, 2227-2230, 2131-2134 and 2325-2328), whose names are also checked. A fighter name is included only when all its damaging abilities have the same normalized mix. It retains mixed profiles, omits zero-damage items, and excludes conflicting duplicate names rather than guessing. Publication status is not a filter. The runtime normalizes the four amounts into fractions once and shares the resulting case-insensitive Map. Only a matched drone or fighter may use the attacker's name as a fallback; other matches require the logged weapon name.

Damage mix is an estimate weighted by logged hit amounts. Logs lack the ship's resistance values and often omit loaded ammo. Unknown damage remains explicitly unknown; NPC base damage attributes are not used to guess combined turret/missile attacks. Mixed-ammo fractions do not account for differing resistances against each type.

To refresh from the same extracted reference directory used above, also extract `dogma_attributes.json`:

```powershell
& 'C:\Program Files\7-Zip\7z.exe' e (Join-Path $npcDataDirectory 'reference-data.tar') dogma_attributes.json "-o$npcDataDirectory" -y
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Update-DamageProfiles.ps1 -ReferenceDirectory $npcDataDirectory
& 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe' /ErrorStdOut .\tests\TestRunner.ahk --headless
```

As with the NPC generator, regeneration loads the large source JSON into memory; this development-time cost does not occur in the application. The generated file records the source build and SHA256. Commit only the generator and generated output, not downloaded data.
