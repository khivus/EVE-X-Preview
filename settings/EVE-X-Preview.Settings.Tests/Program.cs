using EveX.Settings.Bridge;
using EveX.Settings.State;
using System.Text.Json.Nodes;

internal static class Program
{
    private static int passed;
    private static async Task<int> Main(string[] args)
    {
        try
        {
            if (args.Length == 2 && args[0] == "--integration") await IntegrationAsync(Path.GetFullPath(args[1]));
            else if (args.Length == 2 && args[0] == "--inspect-host") await InspectHostAsync(Path.GetFullPath(args[1]));
            else if (args.Length == 2 && args[0] == "--retry-integration") await RetryIntegrationAsync(Path.GetFullPath(args[1]));
            else if (args.Length == 2 && args[0] == "--handoff-integration") await HandoffIntegrationAsync(Path.GetFullPath(args[1]));
            else if (args.Length == 0) { ProtocolTests(); ClipboardTests(); ProblemTests(); await ActionQueueTestsAsync(); await StartupReadinessTestAsync();await StateTestsAsync();await ProfileMutationIsolationAsync();await ReconnectRetriesAsync(); }
            else throw new ArgumentException("Use no arguments for unit tests or --integration <isolated installation directory>.");
            Console.WriteLine($"PASS: {passed} assertions."); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine("FAIL: " + error); return 1; }
    }
    private static async Task ActionQueueTestsAsync()
    {
        var queue = new EditorActionQueue();
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var saved = new List<string>();
        var first = queue.RunAsync(async () => { await release.Task; saved.Add("layout:on"); });
        var second = queue.RunAsync(() => { saved.Add("appearance:on"); return Task.CompletedTask; });
        var third = queue.RunAsync(() => { saved.Add("layout:off"); return Task.CompletedTask; });
        Check(queue.PendingCount == 3 && saved.Count == 0, "rapid inheritance clicks remain queued while the first bridge reply is delayed");
        release.SetResult();
        await Task.WhenAll(first, second, third);
        Check(saved.SequenceEqual(new[] { "layout:on", "appearance:on", "layout:off" }) && queue.PendingCount == 0,
            "queued clicks save in order, including reversing an earlier change, before the final refresh");

        release = new(TaskCreationOptions.RunContinuationsAsynchronously);
        var failing = queue.RunAsync(async () => { await release.Task; throw new InvalidOperationException("Rejected edit"); });
        var following = queue.RunAsync(() => { saved.Add("events:on"); return Task.CompletedTask; });
        release.SetResult();
        await ThrowsAsync<InvalidOperationException>(() => failing, "a failed queued edit reports its error");
        await following;
        Check(saved.Last() == "events:on" && queue.PendingCount == 0, "a failed edit does not drop subsequent clicks or leave the queue busy");
    }
    private static async Task ReconnectRetriesAsync() {
        var transport=new ScriptedTransport();await using var editor=new EditorSession(transport);
        await editor.ConnectAsync();await editor.MutateAsync("select-profile",new(){["name"]="Other"});
        editor.Stage("Other|appearance|value","Retained value","set",new(){["id"]="value",["value"]="pending"});
        transport.MetadataFailures=2;var attempts=new List<int>();
        await editor.ReconnectAsync(attempt=>{attempts.Add(attempt);return Task.CompletedTask;});
        Check(attempts.SequenceEqual(new[]{1,2,3}) && editor.Connected && editor.Profile=="Other" && editor.Pending.Count==1,"automatic reconnect retries transient failures and retains profile and unfinished input");
        transport.MetadataFailures=5;attempts.Clear();
        await ThrowsAsync<BridgeException>(()=>editor.ReconnectAsync(attempt=>{attempts.Add(attempt);return Task.CompletedTask;}),"automatic reconnect stops after five attempts");
        Check(attempts.Count==5 && !editor.Connected && editor.Pending.Count==1,"exhausted reconnect preserves unfinished input for a manual retry");
        transport.MetadataFailures=0;await editor.ReconnectAsync();
        Check(editor.Connected,"manual retry can reconnect after the automatic budget is exhausted");
        await editor.DiscardAsync();await editor.MutateAsync("select-profile",new(){["name"]="Other"});
        transport.Restart=true;transport.ResetProfileOnBegin=true;await editor.ApplyAsync();attempts.Clear();
        await editor.ReconnectAsync(attempt=>{attempts.Add(attempt);if(attempt==3)transport.Pid++;return Task.CompletedTask;});
        Check(attempts.Count==3 && editor.Connected && editor.Profile=="Other" && !transport.RecoveredOnLastBegin,"Apply reconnect waits for the new runtime and retains the edited profile without replaying the saved draft");
    }
    private static void Check(bool condition, string label)
    {
        if (!condition) throw new Exception(label);
        passed++; Console.WriteLine("PASS " + label);
    }
    private static async Task ProfileMutationIsolationAsync() {
        var gate=new GatedTransport();await using var editor=new EditorSession(gate);await editor.ConnectAsync();
        await editor.MutateAsync("select-profile",new(){["name"]="Other"});
        editor.Stage("Default|mouse-actions|rows","Gestures","set",new(){["id"]="opaque-field",["value"]="saved"});
        gate.Block=true;var accepting=editor.AcceptPendingAsync();await gate.Entered.Task;
        editor.Stage("Other|appearance|next","Next edit","set",new(){["id"]="opaque-field",["value"]="next"});
        Check(editor.Pending.Single(e=>e.Key.StartsWith("Other|")).Profile=="Other","pending edits keep their rendered profile during a temporary Default mutation");
        var selecting=editor.MutateAsync("select-profile",new(){["name"]="Other"});gate.Release.TrySetResult();
        await accepting;await selecting;
        Check(gate.Inner.LastSetProfile=="Default" && editor.Profile=="Other","profile navigation cannot interleave with accepting and restoring an edit's original profile");
        await editor.DiscardAsync();await editor.EndAsync(false);
    }
    private static async Task StartupReadinessTestAsync() {
        var transport=new ScriptedTransport {NotReadyHellos=2};await using var editor=new EditorSession(transport);
        await editor.ConnectAsync();
        Check(editor.Connected && transport.HelloCount==3,"cold startup waits for host readiness without requiring Reload");
        Check(transport.BeginCount==1,"startup readiness retries acquire the editing lease only once");
        await editor.EndAsync(false);
    }
    private static void Throws<T>(Action action, string label) where T : Exception
    {
        try { action(); } catch (T) { Check(true, label); return; }
        throw new Exception("Expected " + typeof(T).Name + ": " + label);
    }
    private static async Task ThrowsAsync<T>(Func<Task> action, string label) where T : Exception
    {
        try { await action(); } catch (T) { Check(true, label); return; }
        throw new Exception("Expected " + typeof(T).Name + ": " + label);
    }
    private static void ProblemTests() {
        var metadata=new Metadata(new(){["fields"]=new JsonArray(new JsonObject {["id"]="width",["label"]="Thumbnail width",["section"]="layout"}),["events"]=new JsonArray(new JsonObject {["id"]="attack",["label"]="Player attack"})});
        var validation=new JsonObject {["issues"]=new JsonArray(
            new JsonObject {["id"]="F1",["kind"]="shortcut-conflict",["message"]="Shortcut overlaps: next, previous",["severity"]="warning"},
            new JsonObject {["id"]="width",["profile"]="Default",["message"]="Too small",["severity"]="error"},
            new JsonObject {["id"]="attack:soundVolume",["message"]="Out of range",["severity"]="error"})};
        var issues=SettingProblems.Build([],validation,metadata);
        Check(issues.Count==3 && issues.Last().Shortcut && issues.First().Severity=="error","shortcut overlaps are always the last problem");
        Check(issues.Last().Severity=="warning" && issues.Last().Message.Contains("You can save") && issues.Last().Message.Contains("not guaranteed"),"shortcut conflict warning explains permitted saving and uncertain behavior");
        Check(issues[0].Section=="layout" && issues[0].Label=="Thumbnail width" && issues[1].Section=="sounds","problem links resolve scalar and sound assignment categories");
        var pending=new PendingEdit("Default|appearance|color","Default","Label color","set",new(){["id"]="color"},"Invalid color");
        var local=SettingProblems.Build([pending],new(),metadata).Single();
        Check(local.Section=="appearance" && local.Message=="Invalid color","unfinished invalid text is listed with its original category");
        Check(SettingProblems.Build([],new(),metadata).Count==0,"valid settings have no active problems");
        var overlaps=SettingProblems.Build([],new(){["issues"]=new JsonArray(
            new JsonObject {["id"]="character:1",["kind"]="shortcut-conflict",["profile"]="Other",["message"]="Shortcut overlaps: Character Bob"},
            new JsonObject {["id"]="group:Fleet:ForwardsHotkey",["kind"]="shortcut-conflict",["profile"]="Other",["message"]="Shortcut overlaps: Group Fleet"})},metadata);
        Check(overlaps.Count==2 && overlaps[0].Section=="character-shortcuts" && overlaps[1].Section=="character-groups","shortcut overlaps link to their actual character or group category");
    }
    private static void ProtocolTests()
    {
        Check(InstallationIdentity.Hash("C:\\Example\\EveX\\") == InstallationIdentity.Hash("c:\\example\\evex"), "installation case and trailing separator normalization");
        Check(InstallationIdentity.Hash("C:\\Example\\EveX") != InstallationIdentity.Hash("C:\\Example\\Other"), "different installations have different endpoints");
        Check(InstallationIdentity.Hash("C:\\ÉVE\\日本語").Length == 64, "non-ASCII installation hash");
        var request = Protocol.Request("hello", ""); var id = request.Text("id");
        Check(id.Length == 32 && id.All(c => c is >= '0' and <= '9' or >= 'a' and <= 'f'), "bounded lowercase request identifiers");
        var response = new JsonObject { ["version"] = 1, ["id"] = id, ["ok"] = true, ["data"] = new JsonObject { ["message"] = "日本語 · ÉVE" } };
        Check(Protocol.Response(response.ToJsonString(), id).Text("message") == "日本語 · ÉVE", "Unicode response payload");
        response["ok"] = 1; Check(Protocol.Response(response.ToJsonString(), id).Text("message") != "", "numeric AHK true compatibility");
        Throws<BridgeException>(() => Protocol.Response(response.ToJsonString(), "bad"), "mismatched response cannot satisfy a request");
        response["version"] = 2; Throws<BridgeException>(() => Protocol.Response(response.ToJsonString(), id), "incompatible response version");
        Throws<BridgeException>(() => Protocol.Response("{", id), "truncated response rejects");
        Throws<BridgeException>(() => Protocol.Response(new string('x', Protocol.MaximumBytes + 1), id), "oversized response rejects before parsing");
        Throws<BridgeException>(() => Protocol.VerifyHello(new() { ["bridgeVersion"] = 1, ["schemaVersion"] = 2, ["ready"] = true }), "schema compatibility handshake");
        Throws<BridgeException>(() => Protocol.VerifyHello(new() { ["bridgeVersion"] = 1, ["schemaVersion"] = 1, ["ready"] = false }), "startup endpoint not-ready protection");
        response["version"] = 1; response["ok"] = false; response["error"] = new JsonObject { ["code"] = "save-failed", ["message"] = "Disk full", ["details"] = new JsonArray("file") };
        try { Protocol.Response(response.ToJsonString(), id); } catch (BridgeException error) { Check(error.Code == "save-failed" && error.Message == "Disk full" && error.Details is JsonArray, "structured server errors preserve code message details"); }
        var metadata = new Metadata(new JsonObject { ["fields"] = new JsonArray(new JsonObject { ["id"] = "new-setting", ["label"] = "New canonical field", ["kind"] = "bool" }), ["sections"] = new JsonArray(new JsonObject { ["id"] = "new-section", ["fields"] = new JsonArray("new-setting") }) });
        Check(metadata.SectionFields(metadata.Sections.Single()).Single().Text("id") == "new-setting", "unknown ordinary field discovered from metadata without C# schema");
    }
    private static void ClipboardTests()
    {
        var rows = new JsonArray(new JsonArray("ÉVE", "tabs\tand\nnewlines", "\"quoted\""), new JsonArray("日本語", "", "last"));
        Check(JsonNode.DeepEquals(TableClipboard.Read(TableClipboard.Write(rows), 3), rows), "clipboard quoted cells Unicode and blank cells roundtrip");
        Check(TableClipboard.Read("a\tb\r\nc\td\r\n", 2).Count == 2, "clipboard trailing newline does not create an empty row");
        Throws<FormatException>(() => TableClipboard.Read("one\ttwo", 3), "clipboard validates column count");
        Throws<FormatException>(() => TableClipboard.Read("\"unterminated", 1), "clipboard rejects unterminated quotes");
        Throws<FormatException>(() => TableClipboard.Read("\"closed\"junk", 1), "clipboard rejects ambiguous quote suffix");
        var names = TableClipboard.ReadGrid("Alice\r\nБоб\n\nCarol\r\n", 4);
        Check(names.Count == 4 && names.OfType<JsonArray>().All(row => row.Count == 1), "chat names paste as single-column rows into wider tables");
        Check(TableClipboard.ReadGrid("Alice\tFFFFFF\nBob\t000000", 4).OfType<JsonArray>().All(row => row.Count == 2), "partial-width paste preserves its actual columns");
        Throws<FormatException>(() => TableClipboard.ReadGrid("a\tb\tc", 2), "grid paste rejects excess columns");
        var brightness = new ColorBrightness(); brightness.SetRgb(255,128,0);
        Check(brightness.WithValue(0) == ((byte)0,(byte)0,(byte)0), "color value reaches true black");
        Check(brightness.WithValue(10) == ((byte)26,(byte)13,(byte)0), "color value produces dark shades while preserving hue");
        Check(brightness.WithValue(100) == ((byte)255,(byte)128,(byte)0), "color hue survives a round trip through black");
        var neutral = new ColorBrightness(); neutral.SetRgb(0,0,0);
        Check(neutral.WithValue(25) == ((byte)64,(byte)64,(byte)64), "black initial colors can become neutral dark gray");
    }
    private static async Task StateTestsAsync()
    {
        var transport = new ScriptedTransport(); await using var editor = new EditorSession(transport);
        await editor.ConnectAsync(); Check(editor.Connected && editor.Profile == "Default", "client acquires canonical draft lease");
        await editor.SetModalAsync(true); Check(transport.ClientDirty && editor.Dirty, "editing dialog advertises protected work before opening");
        await editor.SetModalAsync(true); await editor.SetModalAsync(false); Check(transport.ClientDirty, "nested dialogs retain protection until outer closes");
        await editor.SetModalAsync(false); Check(!transport.ClientDirty && !editor.Dirty, "cancelled clean dialog releases protection");
        var key = "Default|appearance|opaque-field";
        editor.Stage(key, "Opaque field", "set", new() { ["id"] = "opaque-field", ["value"] = "unfinished invalid text" });
        await editor.ProtectPendingAsync(); Check(editor.Dirty && transport.ClientDirty, "unfinished text advertises a dirty lease");
        transport.RejectSet = true;
        await ThrowsAsync<BridgeException>(() => editor.AcceptPendingAsync(), "canonical validation rejects draft mutation");
        Check(editor.GetPending(key)?.Arguments.Text("value") == "unfinished invalid text" && editor.GetPending(key)?.Error is not null, "invalid input and feedback survive rejection");
        await ThrowsAsync<BridgeException>(() => editor.ApplyAsync(), "pending edits prevent Apply");
        await ThrowsAsync<BridgeException>(() => editor.EndAsync(false), "pending edits prevent clean close");
        await ThrowsAsync<BridgeException>(() => editor.MutateAsync("profile", new() { ["action"] = "rename", ["name"] = "Renamed" }), "profile rename cannot orphan pending edits");
        await editor.MutateAsync("select-profile", new() { ["name"] = "Other" });
        transport.RejectSet = false; await editor.AcceptPendingAsync(); Check(editor.Pending.Count == 0 && !transport.ClientDirty, "accepted input clears pending guard");
        Check(transport.LastSetProfile == "Default" && editor.Profile == "Other", "unfinished edits retain original profile across navigation and restore selected profile");
        var rowKey = "Other|app-groups|rows|app-group-members|Group";
        editor.Stage(rowKey, "Members", "rows", new() { ["kind"] = "app-group-members", ["group"] = "Group", ["rows"] = new JsonArray(new JsonArray("unfinished.exe", "edited title")) });
        editor.AppendPendingRows(rowKey, "Members", "app-group-members", "Group", new JsonArray(new JsonArray("stale.exe", "old title")), new JsonArray(new JsonArray("selected.exe", "window title")));
        Check(editor.GetPending(rowKey)?.Arguments.Array("rows").Count == 2 && editor.GetPending(rowKey)?.Arguments.Array("rows")[0]?[0]?.ToString() == "unfinished.exe", "running-window picker appends without replacing unfinished member rows");
        await editor.AcceptPendingAsync(rowKey);
        var gainKey = "Other|sounds|gain|C:\\Audio\\alert.wav";
        editor.Stage(gainKey, "Gain", "sound", new() { ["action"] = "gain", ["path"] = "C:\\Audio\\alert.wav", ["volume"] = "unfinished" });
        await ThrowsAsync<BridgeException>(() => editor.MutateAsync("sound", new() { ["action"] = "remove", ["path"] = "C:\\Audio\\alert.wav" }), "sound removal cannot orphan unfinished file gain");
        await editor.DiscardPendingAsync(gainKey); Check(editor.GetPending(gainKey) is null, "individual pending edit can be discarded without discarding canonical draft");
        transport.ResetProfileOnBegin = true;
        transport.DisconnectNext = true; await ThrowsAsync<BridgeException>(() => editor.ReadAsync("snapshot"), "connection loss surfaced");
        Check(!editor.Connected, "connection failure disables mutation"); await editor.ConnectAsync(); Check(transport.Recovered, "reconnect restores opaque canonical draft");
        Check(editor.Profile=="Other","reconnect keeps the edited profile when the new runtime starts on Default");
        Check(editor.Validation.Array("issues").Count==0,"reconnect clears stale validation before the new host is validated");
        editor.Stage(key, "Opaque field", "set", new() { ["id"] = "opaque-field", ["value"] = "still unfinished" }); await editor.DiscardAsync(); Check(!editor.Dirty && editor.Pending.Count == 0, "explicit Discard clears canonical and unfinished state");
        await editor.MutateAsync("set", new() { ["id"] = "opaque-field", ["value"] = "saved" });
        transport.FailSave = true; await ThrowsAsync<BridgeException>(() => editor.ApplyAsync(), "failed save surfaced"); Check(editor.Dirty, "failed save preserves editable canonical draft");
        transport.FailSave = false; transport.Restart = true; Check(await editor.ApplyAsync(), "Apply confirms restart");
        var beginCount = transport.BeginCount;
        await ThrowsAsync<BridgeException>(() => editor.ConnectAsync(), "Apply reconnect rejects previous runtime PID during delayed restart");
        Check(transport.BeginCount == beginCount && !editor.Connected, "previous runtime lease is not reacquired while restart remains pending");
        transport.Pid++;
        await editor.ConnectAsync(); Check(!transport.RecoveredOnLastBegin, "confirmed save does not replay the pre-save recovery draft");
        await editor.EndAsync(false); Check(!editor.Connected, "end releases session");
        var lostTransport = new ScriptedTransport { LoseApplyResponse = true, DelayedRecoveryAttempts = 1, Restart = true };
        await using var lostEditor = new EditorSession(lostTransport); await lostEditor.ConnectAsync();
        await lostEditor.MutateAsync("set", new() { ["id"] = "opaque-field", ["value"] = "saved-but-response-lost" });
        await ThrowsAsync<BridgeException>(() => lostEditor.ApplyAsync(), "lost Apply acknowledgement surfaces an uncertain outcome");
        var previousBegin = lostTransport.BeginCount;
        await ThrowsAsync<BridgeException>(() => lostEditor.ConnectAsync(), "unresolved delayed Apply response blocks old draft recovery");
        Check(lostTransport.BeginCount == previousBegin && lostEditor.Dirty, "uncertain Apply neither imports old draft nor silently discards it");
        lostTransport.Pid++;
        await lostEditor.ConnectAsync();
        Check(!lostTransport.RecoveredOnLastBegin && !lostEditor.Dirty && !lostTransport.HasUncertainApply, "delayed confirmed save reconnects new runtime without replaying old operations");
        await lostEditor.EndAsync(false);
    }
    private static async Task IntegrationAsync(string installation)
    {
        Console.WriteLine("Real canonical bridge integration against isolated fixture: " + installation);
        await using var editor = new EditorSession(new BridgeTransport(installation));
        await editor.ConnectAsync(); Check(editor.Connected, "real WM_COPYDATA hello metadata begin");
        var metadata = editor.Metadata!;
        Check(metadata.Fields.Any() && metadata.Sections.Any() && metadata.Tables.Any() && metadata.Events.Any(), "real metadata covers fields sections tables events");
        var orderedGroup = "Ordered shortcuts " + Guid.NewGuid().ToString("N")[..6];
        await editor.MutateAsync("group", new() { ["kind"] = "characters", ["action"] = "add", ["name"] = orderedGroup });
        var groupRows = await editor.ReadAsync("rows", new() { ["kind"] = "group-members", ["group"] = orderedGroup });
        Check(groupRows.Array("groupShortcuts").Select(row => row.Text("key")).SequenceEqual(new[] { "ForwardsHotkey", "BackwardsHotkey", "FirstCharHotkey" }), "character group presentation order is Next Previous First");
        await editor.MutateAsync("group", new() { ["kind"] = "characters", ["action"] = "delete", ["name"] = orderedGroup });
        var original = editor.Snapshot.Object("values").Object("SettingsTheme")["value"]!.DeepClone();
        var customThemes = metadata.Value.Array("themes").OfType<JsonObject>().ToArray();
        var next = customThemes.FirstOrDefault()?.Text("id") ?? (JsonExtensions.Display(original) == "dark" ? "light" : "dark");
        foreach (var palette in customThemes)
        {
            await editor.MutateAsync("set", new() { ["id"] = "SettingsTheme", ["value"] = palette.Text("id") });
            Check(editor.Snapshot.Object("values").Object("SettingsTheme").Text("value") == palette.Text("id"), "canonical theme choice roundtrip: " + palette.Text("label"));
        }
        var key = "Default|interface|SettingsTheme";
        editor.Stage(key, "Theme", "set", new() { ["id"] = "SettingsTheme", ["value"] = "not-a-valid-theme" }); await editor.ProtectPendingAsync();
        Check(editor.Snapshot.Bool("clientDirty"), "real pending edit protects backend lease");
        await ThrowsAsync<BridgeException>(() => editor.AcceptPendingAsync(key), "real canonical validation rejects invalid theme");
        Check(editor.Pending.Count == 1, "real rejected text retained");
        editor.Stage(key, "Theme", "set", new() { ["id"] = "SettingsTheme", ["value"] = next }); await editor.AcceptPendingAsync(key);
        Check(editor.Snapshot.Object("values").Object("SettingsTheme").Text("value") == next, "real scalar roundtrip");
        var name = "Bridge test ÉVE " + Guid.NewGuid().ToString("N")[..6];
        await editor.MutateAsync("profile", new() { ["action"] = "add", ["name"] = name, ["copyCurrent"] = true, ["inheritance"] = editor.Snapshot.Object("creationInheritance").DeepClone() });
        Check(editor.Profile == name, "real Unicode profile creation");
        var inheritable = metadata.Sections.First(s => s.Text("group") != "").Text("id");
        await editor.MutateAsync("inheritance", new() { ["section"] = inheritable, ["enabled"] = true });
        Check(editor.Snapshot.Object("inheritance").Bool(inheritable), "real section inheritance");
        await editor.MutateAsync("profile", new() { ["action"] = "rename", ["name"] = name + " renamed" }); Check(editor.Profile.EndsWith(" renamed"), "real profile rename");
        await editor.MutateAsync("profile", new() { ["action"] = "delete" }); await editor.MutateAsync("undo", new()); Check(editor.Profile.EndsWith(" renamed"), "real profile delete undo");
        await editor.DiscardAsync();
        await editor.MutateAsync("set", new() { ["id"] = "SettingsTheme", ["value"] = next });
        Check((await editor.ValidateAsync()).Bool("canApply"), "real canonical draft validation");
        var restart = await editor.ApplyAsync(); Check(true, "real canonical Apply acknowledgement");
        if (restart) await ReconnectAsync(editor);
        Check(editor.Snapshot.Object("values").Object("SettingsTheme").Text("value") == next, restart ? "real saved value survives runtime restart" : "real saved value survives offline canonical Apply");
        await editor.MutateAsync("set", new() { ["id"] = "SettingsTheme", ["value"] = original });
        if (await editor.ApplyAsync()) await ReconnectAsync(editor);
        await editor.EndAsync(false); Check(!editor.Connected, "real clean end releases lease");
    }
    private static async Task InspectHostAsync(string installation)
    {
        using var transport = new BridgeTransport(installation);
        // The production editor is still performing its initial metadata request.
        // AHK allows one WM_COPYDATA handler at a time; retry this read-only observer.
        for (var attempt = 0; ; attempt++)
        {
            try
            {
                await transport.EnsureAvailableAsync();
                Protocol.VerifyHello(await transport.SendAsync("hello", ""));
                Check(!transport.RuntimeAvailable, "production settings editor started a ready offline canonical host");
                var metadata = new Metadata(await transport.SendAsync("metadata", ""));
                Check(metadata.Fields.Any() && metadata.Sections.Any(), "embedded host exports usable canonical metadata without taking the editor lease");
                return;
            }
            catch (BridgeException error) when (attempt < 20 && error.Code is "disconnected" or "timeout" or "not-ready")
            {
                await Task.Delay(500);
            }
        }
    }
    private static async Task ReconnectAsync(EditorSession editor)
    {
        for (var attempt = 0; attempt < 30; attempt++)
        {
            await Task.Delay(500);
            try { await editor.ConnectAsync(); Check(editor.Connected, "real reconnect after restart"); return; }
            catch (BridgeException error) when (error.IsDisconnected) { }
        }
        throw new TimeoutException("Runtime did not restart/reconnect after Apply.");
    }
    private static async Task RetryIntegrationAsync(string installation)
    {
        await using var editor = new EditorSession(new BridgeTransport(installation));
        await editor.ConnectAsync();
        await editor.MutateAsync("set", new() { ["id"] = "SettingsTheme", ["value"] = "dark" });
        await ThrowsAsync<BridgeException>(() => editor.ApplyAsync(), "real response-write failure leaves Apply outcome uncertain");
        await editor.ConnectAsync();
        Check(editor.Connected && !editor.Dirty, "same-ID retry recovers cached successful Apply");
        Check(editor.Snapshot.Object("values").Object("SettingsTheme").Text("value") == "dark", "recovered saved theme is authoritative");
        Check(File.ReadAllText(Path.Combine(installation, "fixture-save-count.txt")).Trim() == "1", "response retry does not repeat canonical save");
        await editor.EndAsync(false);
    }
    private static async Task HandoffIntegrationAsync(string installation)
    {
        await using var editor = new EditorSession(new BridgeTransport(installation));
        await editor.ConnectAsync();
        Check(editor.Role == "offline", "handoff starts in canonical offline host");
        await editor.MutateAsync("set", new() { ["id"] = "SettingsTheme", ["value"] = "matrix" });
        var key = "Default|appearance|ThumbnailTextSize";
        editor.Stage(key, "Label size", "set", new() { ["id"] = "ThumbnailTextSize", ["value"] = "unfinished" });
        await editor.ProtectPendingAsync();
        var compiled = Path.Combine(installation, "EVE-X-Preview.exe");
        var start = new System.Diagnostics.ProcessStartInfo(File.Exists(compiled) ? compiled : Path.Combine(installation, "AutoHotkey64.exe")) { UseShellExecute = false, WorkingDirectory = installation, CreateNoWindow = true };
        if (!File.Exists(compiled)) start.ArgumentList.Add(Path.Combine(installation, "Main.ahk"));
        using var runtime = System.Diagnostics.Process.Start(start) ?? throw new Exception("Could not start isolated handoff runtime.");
        try
        {
            for (var attempt = 0; attempt < 30; attempt++)
            {
                try { await editor.ConnectAsync(); if (editor.Role == "runtime") break; }
                catch (BridgeException error) when (error.IsDisconnected) { }
                await Task.Delay(200);
            }
            Check(editor.Role == "runtime", "reconnect hands offline session to runtime");
            Check(editor.Snapshot.Object("values").Object("SettingsTheme").Text("value") == "matrix", "handoff retains accepted canonical draft");
            Check(editor.GetPending(key)?.Arguments.Text("value") == "unfinished", "handoff retains unfinished local text");
            await ThrowsAsync<BridgeException>(() => editor.EndAsync(false), "handoff still protects unsaved changes");
            await editor.DiscardAsync(); await editor.EndAsync(false);
        }
        finally { if (!runtime.HasExited) { runtime.Kill(); await runtime.WaitForExitAsync(); } }
    }
}

internal sealed class ScriptedTransport : IBridgeTransport
{
    public bool RuntimeAvailable => true;
    public bool ClientDirty, RejectSet, DisconnectNext, Recovered, RecoveredOnLastBegin, FailSave, Restart;
    private bool dirty;
    private string profile = "Default";
    public string LastSetProfile = "";
    public int Pid = 501, BeginCount;
    public int NotReadyHellos,HelloCount,MetadataFailures;
    public bool LoseApplyResponse;
    public bool ResetProfileOnBegin;
    public int DelayedRecoveryAttempts;
    private JsonObject? retainedApply;
    public bool HasUncertainApply => retainedApply is not null;
    public Task<JsonObject?> RecoverApplyAsync(CancellationToken cancellation = default)
    {
        if (DelayedRecoveryAttempts-- > 0) return Task.FromResult<JsonObject?>(null);
        var outcome = retainedApply; retainedApply = null; return Task.FromResult(outcome);
    }
    public Task EnsureAvailableAsync(CancellationToken cancellation = default) => Task.CompletedTask;
    private JsonObject Snapshot() => new() { ["session"] = "fixture-session", ["profile"] = profile, ["profiles"] = new JsonArray(new JsonObject {["name"]="Default"},new JsonObject {["name"]="Other"}), ["dirty"] = dirty || ClientDirty, ["clientDirty"] = ClientDirty, ["draft"] = new JsonObject { ["opaque"] = "canonical-only", ["dirty"] = dirty } };
    public Task<JsonObject> SendAsync(string method, string session, JsonObject? arguments = null, CancellationToken cancellation = default)
    {
        if (DisconnectNext) { DisconnectNext = false; throw new BridgeException("timeout", "Fixture disconnect"); }
        switch (method)
        {
            case "hello": HelloCount++;return Task.FromResult(new JsonObject { ["bridgeVersion"] = 1, ["schemaVersion"] = 1, ["ready"] = NotReadyHellos--<=0, ["role"] = "runtime", ["pid"] = Pid });
            case "metadata": if(MetadataFailures-->0)throw new BridgeException("timeout","Transient metadata failure");return Task.FromResult(new JsonObject { ["bridgeVersion"] = 1, ["schemaVersion"] = 1 });
            case "begin": BeginCount++;if(ResetProfileOnBegin)profile="Default"; RecoveredOnLastBegin = arguments?["recovery"] is not null; Recovered |= RecoveredOnLastBegin; break;
            case "client-state": ClientDirty = arguments.Bool("dirty"); break;
            case "select-profile": profile = arguments.Text("name"); break;
            case "set": if (RejectSet) throw new BridgeException("validation", "Invalid input"); LastSetProfile = profile; dirty = true; break;
            case "validate": return Task.FromResult(new JsonObject { ["canApply"] = true, ["issues"] = new JsonArray(), ["conflicts"] = new JsonArray() });
            case "apply":
                if (FailSave) throw new BridgeException("save-failed", "Disk full");
                dirty = false;
                var applied = new JsonObject { ["applied"] = true, ["restartRequired"] = Restart, ["snapshot"] = Snapshot() };
                if (LoseApplyResponse) { retainedApply = applied; throw new BridgeException("timeout", "Save completed but acknowledgement was delayed."); }
                return Task.FromResult(applied);
            case "discard": dirty = false; ClientDirty = false; break;
            case "end": return Task.FromResult(new JsonObject { ["closed"] = true });
        }
        return Task.FromResult(Snapshot());
    }
    public void Dispose() { }
}

internal sealed class GatedTransport : IBridgeTransport {
    public ScriptedTransport Inner {get;}=new();
    public bool RuntimeAvailable=>true;
    public bool Block;
    public TaskCompletionSource Entered {get;}=new(TaskCreationOptions.RunContinuationsAsynchronously);
    public TaskCompletionSource Release {get;}=new(TaskCreationOptions.RunContinuationsAsynchronously);
    public Task EnsureAvailableAsync(CancellationToken cancellation=default)=>Task.CompletedTask;
    public async Task<JsonObject> SendAsync(string method,string session,JsonObject? arguments=null,CancellationToken cancellation=default) {
        if(Block && method=="set") {Block=false;Entered.TrySetResult();await Release.Task;}
        return await Inner.SendAsync(method,session,arguments,cancellation);
    }
    public void Dispose()=>Inner.Dispose();
}
