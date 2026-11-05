using System.Security.Principal;

namespace EveX.Settings.Bridge;

public sealed class SingleInstance : IDisposable
{
    private readonly Mutex mutex;
    private readonly EventWaitHandle activation;
    private readonly CancellationTokenSource stop = new();
    private readonly Task? listener;
    public bool IsPrimary { get; }
    public event Action? ActivationRequested;
    public SingleInstance(string installation)
    {
        var user = WindowsIdentity.GetCurrent().User?.Value ?? Environment.UserName;
        var name = "Local\\EVE-X-Preview.Settings/1/" + user + "/" + InstallationIdentity.Hash(installation);
        mutex = new Mutex(true, name, out var created);
        IsPrimary = created;
        activation = new EventWaitHandle(false, EventResetMode.AutoReset, name + "/activate");
        if (IsPrimary) listener = Task.Run(() =>
        {
            while (!stop.IsCancellationRequested)
                if (WaitHandle.WaitAny([activation, stop.Token.WaitHandle]) == 0) ActivationRequested?.Invoke();
        });
    }
    public void ActivateExisting() => activation.Set();
    public void Dispose()
    {
        stop.Cancel(); listener?.Wait(1000); activation.Dispose();
        if (IsPrimary) mutex.ReleaseMutex();
        mutex.Dispose(); stop.Dispose();
    }
}
