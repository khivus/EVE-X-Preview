# Rapid switching acceptance

Run executable/UI checks only inside a disposable Windows Sandbox. This test
injects keyboard and mouse events, starts multiple synthetic EVE windows, and
writes configuration in its checkout. It refuses to run outside `C:\Checks`
under the Sandbox account. Never point it at an installed application.

Use a Sandbox with networking and clipboard disabled, read-only tool/source
inputs, and a writable `C:\Output` mapping for results. **Stop and disable the
guest's `Audiosrv` service before starting any apps.** Do not rely on an
`AudioOutput` XML element to mute Sandbox. Host audio settings need not change.

Keep the remote connection stable during a held-modifier run. Reconnecting the
Sandbox can synchronize host keyboard state into the guest and release a
synthetically held Shift. This was reproduced with the main app absent: the
guest received untagged left/right Shift-up events during a connection refresh.
Avoid host Shift input during the short held-modifier acceptance check and do
not reconnect until it finishes. Such a reset invalidates the run; it is not
evidence that the application released the modifier.

**Never suspend the Sandbox's host-side remote-viewer process to isolate input.**
That process can own host keyboard hooks; suspending it can interfere with host
shortcuts. Do not filter key releases or re-press lost modifiers to make a run
pass. The guest driver now tags its modifier events and stops immediately if
held Shift is released before a switch or command, leaving the failed run's
logs intact for investigation.

Copy this checkout beneath `C:\Checks`, and stage AutoHotkey v2 x64 plus Ahk2Exe
in the guest. From a guest PowerShell:

```powershell
Stop-Service Audiosrv -Force
Set-Service Audiosrv -StartupType Disabled
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Checks\candidate\tools\Test-RapidSwitchGuest.ps1 `
  -Runtime C:\Runtime\AutoHotkey64.exe `
  -Compiler C:\Input\ahk-tools\Compiler\Ahk2Exe.exe `
  -Label candidate -Clients 16
```

The helper compiles the actual main app and a C# Windows Forms receiver using
the guest's inbox .NET Framework x64 compiler. It creates 16 independent
receiver processes named `exefile.exe`. Game DVR startup discovery is disabled only in
the guest to prevent its missing `ms-gamingoverlay` handler stealing focus.
The driver registers `*F13` as Character-group Next through the real app config.
It tags each injected F down/up so the analyzer can check the receiving client,
Shift state, duplicates, missing events and split down/up pairs.

Before switching, a standalone receiver regression sends 32 back-to-back
Shift/F commands with a 25 ms delay on every receiver log write. The app is
absent during this check. Both tagged F events and their Shift state must
survive the slow logging before the helper starts the main acceptance test.
These results are saved under the `receiver` subdirectory.

The preflight also releases Shift deliberately and sends one F pair. The same
analyzer must reject that negative control specifically for missing Shift while
confirming both F events reached the intended client. Its expected rejection is
saved in `modifier-release-analysis.txt`; the guest helper treats it as a
passing self-check only when the other delivery checks succeeded.

The receiver records the native message queue serially, without pumping
messages during log writes. The previous AHK receiver used `OnMessage`, whose
single-thread callback limit can skip observation of a second WM_KEYDOWN while
the Shift-down callback is still logging. A delayed-logging reproduction with
that receiver missed all 32 F key-downs even with the main app absent. The
replacement avoids that observer failure; the analyzer still rejects missing,
duplicate, wrong-client, unshifted, and split F events. See the
[AutoHotkey OnMessage documentation](https://www.autohotkey.com/docs/v2/lib/OnMessage.htm#Critical)
for the callback scheduling restriction.

The primary settings match the reported configuration: preserved shortcuts on
logout, remembered group positions, 30 ms group-switch minimum, three focus
checks at 25 ms, 50 ms window polling, and startup staggering/DWM pause/layout
retry disabled. A secondary case disables remembered positions. Both modes
also exercise manual selection followed by Next, preserving their distinct
existing behavior. Each keyboard case starts a fresh application and primes
the group cursor at member 1; expected recipients follow the configured order.

Required cases cover Shift held continuously, 10 ms and 40 ms F taps, 10 ms and
30 ms F13-to-F gaps, and separately pressing Shift for each command. Two stress
cases record zero-gap input and 20 ms switch intervals, the latter intentionally
below the configured 30 ms limit. Stress results are reported but excluded from
the pass/fail gate; foreground scheduling and the configured hold limit do not
guarantee arbitrary-speed delivery. Requested millisecond timings are subject
to Windows scheduling and timer resolution.

Results include `driver.csv`, per-client event logs, the exact settings,
`manual.csv`, `summary.csv`, `delivery.csv`, and `exit-code.txt`. To recheck saved
results without starting any applications:

```powershell
.\tools\Measure-RapidSwitch.ps1 -Directory C:\Output\candidate -AssertAll
```

For a before/after reproduction, use fresh source copies of the base and fix
with the same acceptance files. Run sequentially in the same muted guest;
concurrent desktop suites would invalidate focus measurements. The test is
standalone and is not included in `tests/TestRunner.ahk`. Run that existing
suite separately in Sandbox for broader regression coverage.

These receivers test Windows input routing, not EVE's rendering, gameplay
command processing, or live-account behavior. The timing fix only removes the
delay on the application's internal foreground signal and restores the
caller's send settings. It does not buffer, replay, or broadcast gameplay keys.

## Initial timing validation

The compiled application was tested in a muted Windows Sandbox with AutoHotkey
2.0.19 and 16 independent receivers, comparing dev `38c65ee` with this change.
The final run produced the following results:

| Case | Commands | Wrong recipient before | Wrong recipient after |
| --- | ---: | ---: | ---: |
| Held Shift, 10 ms gap, 10 ms F tap | 64 | 34 | 0 |
| Held Shift, 10 ms gap, 40 ms F tap | 64 | 39 | 0 |
| Held Shift, 30 ms gap | 64 | 0 | 0 |
| Shift released between commands | 64 | 37 | 0 |
| Remembered positions off, held Shift, 10 ms gap | 64 | 59 | 0 |

All 320 required deliveries after the fix retained Shift, delivered both key
events to the same intended client, and had no missing or duplicate events.
All 64 manual-click checks passed after the fix. The baseline's manual-click
phase did not finish because one synthetic window failed to receive the click;
the table compares its completed keyboard trials only.

Stress cases remain outside the guarantee: with a zero gap, 1 of 64 commands
went to the wrong recipient; at a 20 ms switch period below the configured
30 ms hold limit, 48 of 64 did so. Neither case is counted as a required pass.

The broader AHK suite completed with 410 passed and 9 failed. All nine failures
require native audio playback with an available audio device. With `Audiosrv`
stopped and disabled, those same nine cases also failed on unmodified dev with
the same errors. Audio playback was not validated, and the service remained
disabled throughout this comparison.

## Receiver follow-up validation

A subsequent run of the original AHK observer reported one missing F key-down
in 320 required commands. The correct client was foreground and recorded its
Shift-down and F-up. The isolated slow-logging reproduction described above
confirmed that the observer could miss key events without the app running.

After replacing that observer, two independent compiled 16-client runs each
passed all 320 required commands and all 64 manual-click checks. Both also
passed the 32-command delayed-logging regression: 640/640 required routing
checks, 128/128 manual checks, and 64/64 slow-log commands in total, with no
missing, duplicate, unshifted, wrong-recipient, or split required events.
The full suite again finished with 410 passed and the same 9 audio-dependent
failures, matching the muted baseline's failure names and messages exactly.

The stress cases still expose routing limits: zero-gap input misrouted 1/64
and 4/64 commands in the two runs; the 20 ms switch period misrouted 49/64 and
59/64. None of the tagged down/up events were missing from the native receiver
logs, including these stress cases. Missing-event assertions remain enabled.

## Modifier-reset follow-up

A connection-refresh probe reproduced the unexpected Shift release with the
main app absent. The safe follow-up uses a stable, normally running connection
and a brief keyboard-free test window. No host process is paused or blocked.

The prepared 16-client run passed all 320 required commands and all 64 manual
checks, with no missing F events, missing Shift, duplicates, wrong recipients,
or split required pairs. The 32-command slow-logging check passed, and the
intentional-release negative control failed only for missing Shift as intended.
Zero-gap stress misrouted 1/64 commands; below-hold-limit stress misrouted 56/64.
Audio remained stopped and disabled throughout.

The preceding full-suite attempt did not complete: 347 tests passed and 10 were
reported failed before an unhandled destroyed-control error at
`src/SettingsTheme.ahk:568`. Nine failures were the known audio-dependent cases;
the other was a sharing violation while `TestRunner.ahk` wrote its results log.
That attempt did not establish full-suite merge readiness. Those failures were
recorded separately from the passing input acceptance.

## Log sharing and paint lifecycle follow-up

Two focused regressions reproduced the failures before the fixes: a reader
opened between log writes prevented the next append with Windows error 32,
and a native `WM_GETFONT` handler destroyed its control before the painter's
next property access.

The test runner now owns one write handle until completion, flushes each line
for live readers, and closes the handle on exit. Incompatible readers cannot
take write access away between lines. Snapshot tools must open their reader
with write sharing (for .NET, `FileAccess.Read` and `FileShare.ReadWrite`,
optionally also `FileShare.Delete`). The Sandbox monitor used an explicit
shared `FileStream` while collecting the full-suite log. Logging errors stay
outside the test-assertion catch so infrastructure failures do not become
misleading assertion failures or count one test as both passed and failed.

Static-control callbacks defer ordinary script interruption during their short
native callback and restore the previous interruption setting afterward. If
a native message synchronously destroys the window, the stale paint is
cancelled; errors on surviving windows still propagate. Drawing-context state
and `BeginPaint`/`EndPaint` cleanup are protected by `finally`.

The focused regressions cover incompatible readers, 100 live Unicode log
snapshots, destroyed controls through direct and callback painting, normal
painting, and propagation of a renderer error with drawing state restored.
Run them inside the muted guest with:

```powershell
C:\Runtime\AutoHotkey64.exe /ErrorStdOut C:\Checks\candidate\tests\TestRunner.ahk --headless --filter RunnerPaint
```

Validation against `dev` at `38c65ee` completed in a fresh, muted Sandbox:

- Before fixes: both targeted regressions failed with the original errors.
- After fixes: all five focused regressions passed.
- Full suite: **415 passed, 9 failed, 424 total**, process exit 1, with no
  unhandled errors. The nine failure names and messages exactly match the
  previously captured muted-dev audio baseline. The previously affected input
  pulse slider and sidebar dragging tests both passed.
- Main application and updater compiled successfully inside the guest.
- `Audiosrv` remained stopped and disabled, including the final build check.
  No host process or keyboard input was paused or blocked.

The 16-client input results above are from the preceding acceptance run; this
follow-up reran the full suite and the log/paint regressions. Audio playback
remains unvalidated, and the audio-dependent failures remain reported.
