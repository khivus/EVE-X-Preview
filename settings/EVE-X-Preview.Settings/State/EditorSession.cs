using System.Text.Json.Nodes;
using EveX.Settings.Bridge;

namespace EveX.Settings.State;

public sealed record PendingEdit(string Key, string Profile, string Label, string Method, JsonObject Arguments, string? Error = null);

// The client owns presentation state and unfinished text only. All settings
// parsing, inheritance, merging, validation and persistence remain in AHK.
public sealed class EditorSession(IBridgeTransport transport) : IAsyncDisposable
{
    private readonly SemaphoreSlim operations = new(1, 1);
    private readonly Dictionary<string, PendingEdit> pending = new(StringComparer.Ordinal);
    private JsonNode? recovery;
    private bool appliedAwaitingReconnect;
    private DateTime restartBegan;
    private int runtimePid;
    private int restartingPid;
    private int modalCount;
    private bool skipRecoveryAfterConfirmedApply;
    public Metadata? Metadata { get; private set; }
    public JsonObject Snapshot { get; private set; } = new();
    public JsonObject Validation { get; private set; } = new();
    public bool Connected { get; private set; }
    public string Role { get; private set; } = "";
    public string SessionId => Snapshot.Text("session");
    public string Profile => Snapshot.Text("profile", "Default");
    public bool Dirty => Snapshot.Bool("dirty") || HasLocalProtectedWork;
    public int ChangeCount => Snapshot.Int("changeCount") + pending.Values.Count(edit => !PendingAlreadyCounted(edit));
    public bool HasActiveModal => modalCount != 0;
    private bool HasLocalProtectedWork => pending.Count != 0 || modalCount != 0;
    public IReadOnlyCollection<PendingEdit> Pending => pending.Values;
    public event Action? Changed;
    public string? Error { get; private set; }

    public async Task ConnectAsync(CancellationToken cancellation = default)
    {
        await operations.WaitAsync(cancellation);
        try
        {
            var selectedProfile = Snapshot.Text("profile");
            if (transport.HasUncertainApply)
            {
                var outcome = await transport.RecoverApplyAsync(cancellation);
                if (outcome is null) throw new BridgeException("uncertain-apply", "The previous save outcome is still unconfirmed. Your draft is retained without replaying it. Reconnect retries the saved acknowledgement.");
                if (!outcome.Bool("applied")) throw new BridgeException("save-failed", "The retained save response did not confirm a successful Apply. Your draft is retained.");
                if (outcome["snapshot"] is JsonObject appliedSnapshot) Accept(appliedSnapshot);
                skipRecoveryAfterConfirmedApply = true;
                appliedAwaitingReconnect = outcome.Bool("restartRequired");
                restartBegan = DateTime.UtcNow; restartingPid = runtimePid;
            }
            if (appliedAwaitingReconnect && !transport.RuntimeAvailable && DateTime.UtcNow - restartBegan < TimeSpan.FromSeconds(12))
                throw new BridgeException("not-ready", "Waiting for EVE-X-Preview to restart after the confirmed save.");
            var hello = await WaitForReadyHostAsync(cancellation);
            if (appliedAwaitingReconnect && (hello.Text("role") != "runtime" || hello.Int("pid") == restartingPid))
                throw new BridgeException("not-ready", "Waiting for the previous EVE-X-Preview process to finish restarting. Your editor state is retained.");
            Role = hello.Text("role");
            if (Role == "runtime") runtimePid = hello.Int("pid");
            var contract = await transport.SendAsync("metadata", "", cancellation: cancellation);
            if (contract.Int("bridgeVersion") != 1 || contract.Int("schemaVersion") != 1)
                throw new BridgeException("incompatible", "The shared settings metadata is incompatible with this editor.");
            Metadata = new Metadata(contract);
            var begin = new JsonObject { ["clientPid"] = Environment.ProcessId };
            if (recovery is not null && !appliedAwaitingReconnect && !skipRecoveryAfterConfirmedApply) begin["recovery"] = recovery.DeepClone();
            Accept(await transport.SendAsync("begin", "", begin, cancellation));
            if (selectedProfile != "" && selectedProfile != Profile && Snapshot.Array("profiles").Any(p => p.Text("name") == selectedProfile))
                Accept(await transport.SendAsync("select-profile", SessionId, new JsonObject { ["name"] = selectedProfile }, cancellation));
            appliedAwaitingReconnect = false;
            skipRecoveryAfterConfirmedApply = false;
            Connected = true; Error = null;
            Validation = new();
            Accept(await transport.SendAsync("client-state", SessionId, new JsonObject { ["dirty"] = HasLocalProtectedWork }, cancellation));
        }
        catch (Exception error) { Connected = false; Error = error.Message; throw; }
        finally { operations.Release(); Changed?.Invoke(); }
    }

    private bool PendingAlreadyCounted(PendingEdit edit)
    {
        var draft = Snapshot.Object("draft");
        if (draft["baseline"] is null) return false;
        JsonNode? baseline = draft["baseline"];
        JsonNode? current = draft["config"];
        string[] path;
        if (edit.Method == "set" && Metadata?.Field(edit.Arguments.Text("id")) is { } field)
            path = (field.Bool("global") ? Array.Empty<string>() : new[] { "_Profiles", edit.Profile }).Concat(field.Array("path").Select(JsonExtensions.Display)).ToArray();
        else if (edit.Method == "event")
            path = ["_Profiles", edit.Profile, "Game Logs Monitoring", "monitoredEvents", edit.Arguments.Text("id"), edit.Arguments.Object("values").First().Key];
        else return false;
        foreach (var key in path) { baseline = (baseline as JsonObject)?[key]; current = (current as JsonObject)?[key]; }
        return !JsonNode.DeepEquals(baseline, current);
    }

    private async Task<JsonObject> WaitForReadyHostAsync(CancellationToken cancellation)
    {
        var startup=System.Diagnostics.Stopwatch.StartNew();
        while(true) {
            try {
                await transport.EnsureAvailableAsync(cancellation);
                var hello=await transport.SendAsync("hello","",cancellation:cancellation);
                Protocol.VerifyHello(hello);return hello;
            }
            catch(BridgeException error) when((error.Code is "not-ready" or "timeout" or "disconnected") && startup.Elapsed<TimeSpan.FromSeconds(25)) {
                await Task.Delay(200,cancellation);
            }
        }
    }

    public PendingEdit? GetPending(string key) => pending.GetValueOrDefault(key);
    public void Stage(string key, string label, string method, JsonObject arguments)
    {
        pending[key] = new PendingEdit(key, Profile, label, method, arguments.Copy());
        Changed?.Invoke();
    }
    public void AppendPendingRows(string key, string label, string kind, string group, JsonArray canonicalRows, JsonArray additions)
    {
        var rows = (JsonArray)(GetPending(key)?.Arguments["rows"] as JsonArray ?? canonicalRows).DeepClone();
        foreach (var addition in additions) rows.Add(addition?.DeepClone());
        Stage(key, label, "rows", new() { ["kind"] = kind, ["group"] = group, ["rows"] = rows });
    }
    public async Task SetModalAsync(bool active)
    {
        modalCount = Math.Max(0, modalCount + (active ? 1 : -1));
        await ProtectPendingAsync(); Changed?.Invoke();
    }
    public async Task DiscardPendingAsync(string key)
    {
        pending.Remove(key); await ProtectPendingAsync(); Changed?.Invoke();
    }
    public async Task ProtectPendingAsync()
    {
        if (!Connected) return;
        await operations.WaitAsync();
        try { Accept(await transport.SendAsync("client-state", SessionId, new JsonObject { ["dirty"] = HasLocalProtectedWork })); }
        catch (Exception error) { HandleError(error); }
        finally { operations.Release(); Changed?.Invoke(); }
    }
    public async Task AcceptPendingAsync(string? key = null)
    {
        var edits = pending.Values.Where(edit => key is null || edit.Key == key).ToArray();
        var selectedProfile = Profile;
        try
        {
            foreach (var edit in edits)
            {
                try
                {
                    if (Profile != edit.Profile) await MutateAsync("select-profile", new() { ["name"] = edit.Profile });
                    await MutateAsync(edit.Method, edit.Arguments);
                    // Do not remove a newer edit made while the server was working.
                    if (pending.GetValueOrDefault(edit.Key) == edit) pending.Remove(edit.Key);
                }
                catch (Exception error) { if (pending.GetValueOrDefault(edit.Key) == edit) pending[edit.Key] = edit with { Error = error.Message }; Changed?.Invoke(); throw; }
            }
        }
        finally
        {
            if (Connected && Profile != selectedProfile) await MutateAsync("select-profile", new() { ["name"] = selectedProfile });
        }
        await ProtectPendingAsync(); Changed?.Invoke();
    }
    public async Task<JsonObject> MutateAsync(string method, JsonObject arguments)
    {
        if (method == "profile" && arguments.Text("action") is "rename" or "delete" && pending.Values.Any(edit => edit.Profile == Profile))
            throw new BridgeException("pending-edit", "Correct or undo invalid fields for this profile before renaming or removing it.");
        if (method == "group" && arguments.Text("action") is "rename" or "delete" && pending.Values.Any(edit => edit.Profile == Profile && (edit.Arguments.Text("group") == arguments.Text("name") || (edit.Method == "group" && edit.Arguments.Text("name") == arguments.Text("name")))))
            throw new BridgeException("pending-edit", "Correct or undo invalid group member and shortcut fields before renaming or removing this group.");
        if (method == "sound" && arguments.Text("action") == "remove" && pending.Values.Any(edit => edit.Method == "sound" && edit.Arguments.Text("path").Equals(arguments.Text("path"), StringComparison.OrdinalIgnoreCase)))
            throw new BridgeException("pending-edit", "Correct or undo this sound's invalid gain before removing its library reference.");
        await operations.WaitAsync();
        try
        {
            RequireConnection();
            if (HasLocalProtectedWork) await transport.SendAsync("client-state", SessionId, new JsonObject { ["dirty"] = true });
            var result = await transport.SendAsync(method, SessionId, arguments.Copy());
            if (result["session"] is not null) Accept(result);
            else if (result["snapshot"] is JsonObject snapshot) Accept(snapshot);
            Error = null;
            return result;
        }
        catch (Exception error) { HandleError(error); throw; }
        finally { operations.Release(); Changed?.Invoke(); }
    }
    public async Task<JsonObject> ReadAsync(string method, JsonObject? arguments = null)
    {
        await operations.WaitAsync();
        try { RequireConnection(); return await transport.SendAsync(method, SessionId, arguments?.Copy()); }
        catch (Exception error) { HandleError(error); throw; }
        finally { operations.Release(); Changed?.Invoke(); }
    }
    public async Task<JsonObject> ValidateAsync()
    {
        Validation = await ReadAsync("validate"); Changed?.Invoke(); return Validation;
    }
    public async Task<bool> ApplyAsync()
    {
        if (pending.Count != 0) throw new BridgeException("pending-edit", "Correct the highlighted fields before applying. Their values are retained on their original pages.");
        var validation = await ValidateAsync();
        if (!validation.Bool("canApply")) throw new BridgeException("validation", "Resolve the validation issues or conflicting edits shown below before applying.", validation);
        var result = await MutateAsync("apply", new JsonObject());
        if (!result.Bool("applied")) throw new BridgeException("save-failed", "The application did not confirm that the settings were saved. Your draft is retained.");
        if (result["snapshot"] is JsonObject snapshot) Accept(snapshot);
        var restart = result.Bool("restartRequired");
        appliedAwaitingReconnect = restart;
        if (restart) { restartBegan = DateTime.UtcNow; restartingPid = runtimePid; }
        if (restart) { Connected = false; Snapshot["dirty"] = false; }
        else { await MutateAsync("snapshot", new JsonObject()); await ProtectPendingAsync(); }
        Changed?.Invoke(); return restart;
    }
    public async Task DiscardAsync()
    {
        await MutateAsync("discard", new JsonObject());
        pending.Clear();
        await ProtectPendingAsync();
        await ValidateAsync();Changed?.Invoke();
    }
    public async Task EndAsync(bool discard)
    {
        if (!discard && Dirty) throw new BridgeException("unsaved", "Apply or explicitly discard the settings draft before closing.");
        if (Connected) await MutateAsync("end", new JsonObject { ["discard"] = discard });
        Connected = false;
    }
    private void Accept(JsonObject snapshot)
    {
        Snapshot = snapshot.Copy();
        if (snapshot["draft"] is not null) recovery = snapshot["draft"]!.DeepClone();
    }
    private void RequireConnection() { if (!Connected) throw new BridgeException("disconnected", "The settings service is disconnected. Reconnect to continue; your draft and unfinished edits are retained."); }
    private void HandleError(Exception error)
    {
        Error = error.Message;
        if (error is BridgeException { IsDisconnected: true } || error is IOException) Connected = false;
    }
    public async ValueTask DisposeAsync()
    {
        try { if (Connected && !Dirty) await EndAsync(false); } catch (Exception) { }
        transport.Dispose(); operations.Dispose();
    }
}
