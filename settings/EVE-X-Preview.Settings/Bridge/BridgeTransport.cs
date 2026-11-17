using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.Json.Nodes;

namespace EveX.Settings.Bridge;

public interface IBridgeTransport : IDisposable
{
    bool RuntimeAvailable { get; }
    bool HasUncertainApply => false;
    Task<JsonObject?> RecoverApplyAsync(CancellationToken cancellation = default) => Task.FromResult<JsonObject?>(null);
    Task<JsonObject> SendAsync(string method, string session, JsonObject? arguments = null, CancellationToken cancellation = default);
    Task EnsureAvailableAsync(CancellationToken cancellation = default);
}

public sealed class BridgeTransport : IBridgeTransport
{
    private readonly string installation, identity, responseDirectory;
    private readonly NativeWindow sender = new();
    private readonly SemaphoreSlim serial = new(1, 1);
    private Process? offline;
    private nint endpoint;
    private string? uncertainApplyId;
    private byte[]? uncertainApplyBytes;
    private nint uncertainApplyEndpoint;
    private uint uncertainApplyServerPid;
    private long uncertainApplyServerStart;
    public bool HasUncertainApply => uncertainApplyId is not null;
    public bool RuntimeAvailable => Find("runtime") != 0;
    public BridgeTransport(string installation)
    {
        this.installation = InstallationIdentity.Normalize(installation);
        identity = InstallationIdentity.Hash(installation);
        responseDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EVE-X-Preview", "IPC", identity);
    }
    private nint Find(string role) => Native.FindWindow(null, $"EVE-X-Preview.SettingsBridge/1/{identity}/{role}");
    public async Task EnsureAvailableAsync(CancellationToken cancellation = default)
    {
        endpoint = Find("runtime");
        if (endpoint == 0) endpoint = Find("offline");
        if (endpoint == 0)
        {
            if (offline is null || offline.HasExited) offline = StartOfflineHost();
            var watch = Stopwatch.StartNew();
            while (watch.Elapsed < TimeSpan.FromSeconds(12))
            {
                cancellation.ThrowIfCancellationRequested();
                endpoint = Find("runtime");
                if (endpoint == 0) endpoint = Find("offline");
                if (endpoint != 0) break;
                if (offline.HasExited) throw new BridgeException("disconnected", "The canonical settings host could not start. Open the internal editor or verify the AutoHotkey installation.");
                await Task.Delay(150, cancellation);
            }
        }
        if (endpoint == 0) throw new BridgeException("timeout", "The settings service did not become available within 12 seconds.");
        VerifyServerUser();
    }
    private void VerifyServerUser()
    {
        Native.GetWindowThreadProcessId(endpoint, out var pid);
        var process = Native.OpenProcess(0x1000, false, pid);
        if (process == 0) throw new BridgeException("disconnected", "Windows could not verify the settings service identity.");
        nint token = 0;
        try
        {
            if (!Native.OpenProcessToken(process, 0x0008, out token)) throw new BridgeException("disconnected", "Windows could not verify the settings service user.");
            using var user = new WindowsIdentity(token);
            using var current = WindowsIdentity.GetCurrent();
            if (user.User?.Value != current.User?.Value) throw new BridgeException("disconnected", "The settings service belongs to a different Windows user.");
        }
        finally { if (token != 0) Native.CloseHandle(token); Native.CloseHandle(process); }
    }
    private Process StartOfflineHost()
    {
        ProcessStartInfo start;
        var compiled = Path.Combine(installation, "EVE-X-Preview-SettingsHost.exe");
        var embedded=BundledSettingsHost.Extract();
        if (embedded is not null) start=new ProcessStartInfo(embedded);
        else if (File.Exists(compiled)) start = new ProcessStartInfo(compiled);
        else
        {
            var script = Path.Combine(installation, "SettingsHost.ahk");
            if (!File.Exists(script)) throw new BridgeException("disconnected", "No running application or canonical settings host was found beside this editor.");
            var candidates = new[] { Path.Combine(installation, "AutoHotkey64.exe"), Path.Combine(installation, "AutoHotkey", "AutoHotkey64.exe"), Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "AutoHotkey", "v2", "AutoHotkey64.exe"), Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "AutoHotkey", "v2", "AutoHotkey64.exe") };
            var interpreter = candidates.FirstOrDefault(File.Exists);
            if (interpreter is null) throw new BridgeException("disconnected", "Source mode needs AutoHotkey v2. Place AutoHotkey64.exe beside SettingsHost.ahk or use the portable release package.");
            start = new ProcessStartInfo(interpreter);
            start.ArgumentList.Add("/ErrorStdOut"); start.ArgumentList.Add(script);
        }
        start.ArgumentList.Add("--parent"); start.ArgumentList.Add(Environment.ProcessId.ToString());
        start.ArgumentList.Add("--installation");start.ArgumentList.Add(installation);
        start.UseShellExecute = false; start.CreateNoWindow = true; start.WindowStyle = ProcessWindowStyle.Hidden; start.WorkingDirectory = installation;
        return Process.Start(start) ?? throw new BridgeException("disconnected", "Windows could not start the canonical settings host.");
    }
    public async Task<JsonObject> SendAsync(string method, string session, JsonObject? arguments = null, CancellationToken cancellation = default)
    {
        await serial.WaitAsync(cancellation);
        try
        {
            if (endpoint == 0) throw new BridgeException("disconnected", "The settings service is disconnected. Your draft is retained; reconnect to continue.");
            if (method == "apply" && HasUncertainApply) throw new BridgeException("uncertain-apply", "The previous save is still awaiting acknowledgement. Reconnect to resolve its result before applying again.");
            var request = Protocol.Request(method, session, arguments);
            var bytes = Encoding.UTF8.GetBytes(request.ToJsonString() + '\0');
            if (bytes.Length > Protocol.MaximumBytes) throw new BridgeException("invalid-request", "This settings request exceeds the 4 MiB size limit.");
            var id = request.Text("id");
            if (method == "apply")
            {
                Native.GetWindowThreadProcessId(endpoint, out var serverPid);
                using var server = Process.GetProcessById(checked((int)serverPid));
                uncertainApplyServerStart = server.StartTime.ToUniversalTime().Ticks;
                uncertainApplyServerPid = serverPid;
                uncertainApplyEndpoint = endpoint;
                uncertainApplyBytes = bytes;
                uncertainApplyId = id;
            }
            var responsePath = Path.Combine(responseDirectory, id + ".json");
            return await Task.Run(() =>
            {
                cancellation.ThrowIfCancellationRequested();
                var buffer = Marshal.AllocHGlobal(bytes.Length);
                var acknowledged = false;
                try
                {
                    Marshal.Copy(bytes, 0, buffer, bytes.Length);
                    var copy = new Native.CopyData { Identifier = Protocol.CopyDataIdentifier, ByteCount = bytes.Length, Data = buffer };
                    var timeout=method=="metadata"?15000u:3000u;
                    if (Native.SendMessageTimeout(endpoint, 0x004A, sender.Handle, ref copy, 0x0002 | 0x0020, timeout, out var result) == 0)
                        throw new BridgeException("timeout", $"The settings service did not respond within {timeout/1000} seconds. Your draft is retained. Reconnect before continuing.");
                    if (result != 1) throw new BridgeException("disconnected", "The settings service rejected this connection. Verify that both applications run as the same Windows user and privilege level.");
                    using var file = new FileStream(responsePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    if (file.Length > Protocol.MaximumBytes) throw new BridgeException("invalid-response", "The settings service response exceeds 4 MiB.");
                    using var reader = new StreamReader(file, new UTF8Encoding(false, true));
                    JsonObject response;
                    try { response = Protocol.Response(reader.ReadToEnd(), id); acknowledged = true; }
                    catch (BridgeException error) when (error.FromServer) { acknowledged = true; throw; }
                    if (method == "hello")
                    {
                        Native.GetWindowThreadProcessId(endpoint, out var serverPid);
                        if (response.Text("installation") != identity || response.Int("pid") != serverPid)
                            throw new BridgeException("incompatible", "The settings service identity does not match this installation and process.");
                    }
                    return response;
                }
                catch (IOException error) { throw new BridgeException("disconnected", "The settings response could not be read. " + error.Message); }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                    if (acknowledged)
                    {
                        if (method == "apply") ClearUncertainApply();
                        try { File.Delete(responsePath); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                    }
                }
            }, cancellation);
        }
        finally { serial.Release(); }
    }
    public async Task<JsonObject?> RecoverApplyAsync(CancellationToken cancellation = default)
    {
        await serial.WaitAsync(cancellation);
        try
        {
            if (uncertainApplyId is not { } id) return null;
            var path = Path.Combine(responseDirectory, id + ".json");
            var watch = Stopwatch.StartNew();
            var retried = false;
            while (watch.Elapsed < TimeSpan.FromSeconds(3))
            {
                cancellation.ThrowIfCancellationRequested();
                if (File.Exists(path))
                {
                    var acknowledged = false;
                    try
                    {
                        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                        if (file.Length > Protocol.MaximumBytes) throw new BridgeException("invalid-response", "The retained save response exceeds the bridge size limit.");
                        using var reader = new StreamReader(file, new UTF8Encoding(false, true));
                        try { var result = Protocol.Response(reader.ReadToEnd(), id); acknowledged = true; return result; }
                        catch (BridgeException error) when (error.FromServer) { acknowledged = true; throw; }
                    }
                    finally
                    {
                        if (acknowledged)
                        {
                            ClearUncertainApply();
                            try { File.Delete(path); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                        }
                    }
                }
                // AHK may have committed/cached the result while its response
                // file could not be written. Retry the exact request once on
                // that same live server, which recognizes the cached ID. Never
                // send an old Apply to a restarted runtime or a reused HWND.
                if (!retried && watch.ElapsedMilliseconds >= 150 && CanRetryOriginalApply())
                {
                    retried = true;
                    await Task.Run(() =>
                    {
                        var bytes = uncertainApplyBytes!;
                        var buffer = Marshal.AllocHGlobal(bytes.Length);
                        try
                        {
                            Marshal.Copy(bytes, 0, buffer, bytes.Length);
                            var copy = new Native.CopyData { Identifier = Protocol.CopyDataIdentifier, ByteCount = bytes.Length, Data = buffer };
                            Native.SendMessageTimeout(uncertainApplyEndpoint, 0x004A, sender.Handle, ref copy, 0x0002 | 0x0020, 2000, out _);
                        }
                        finally { Marshal.FreeHGlobal(buffer); }
                    }, cancellation);
                }
                await Task.Delay(100, cancellation);
            }
            return null;
        }
        finally { serial.Release(); }
    }
    private bool CanRetryOriginalApply()
    {
        if (uncertainApplyBytes is null || uncertainApplyEndpoint == 0) return false;
        Native.GetWindowThreadProcessId(uncertainApplyEndpoint, out var pid);
        if (pid != uncertainApplyServerPid || (Find("runtime") != uncertainApplyEndpoint && Find("offline") != uncertainApplyEndpoint)) return false;
        try
        {
            using var server = Process.GetProcessById(checked((int)pid));
            return !server.HasExited && server.StartTime.ToUniversalTime().Ticks == uncertainApplyServerStart;
        }
        catch (Exception error) when (error is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception) { return false; }
    }
    private void ClearUncertainApply()
    {
        uncertainApplyId = null; uncertainApplyBytes = null;
        uncertainApplyEndpoint = 0; uncertainApplyServerPid = 0; uncertainApplyServerStart = 0;
    }
    public void Dispose()
    {
        sender.Dispose();
        // The canonical child watches our PID. Requesting end is performed by the
        // session before disposing transport; never kill another running app.
        offline?.Dispose(); serial.Dispose();
    }
}
