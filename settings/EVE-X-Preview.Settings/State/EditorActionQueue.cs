namespace EveX.Settings.State;

// Keep accepted UI actions in order while the settings bridge is awaiting a reply.
public sealed class EditorActionQueue
{
    private readonly SemaphoreSlim gate = new(1, 1);
    public int PendingCount { get; private set; }

    public async Task RunAsync(Func<Task> action)
    {
        PendingCount++;
        await gate.WaitAsync();
        try { await action(); }
        finally { PendingCount--; gate.Release(); }
    }
}
