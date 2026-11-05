using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Controls.Documents;
using Avalonia.Threading;
using EveX.Settings.Bridge;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    private readonly Dictionary<string, CancellationTokenSource> validationTimers = new();
    private readonly Dictionary<string, Action> fieldFeedback = new();
    private readonly Stack<JsonObject> editHistory = new();
    private readonly Dictionary<string, JsonObject> profileBaselines = new();
    private readonly SemaphoreSlim liveEdits = new(1, 1);

    private async void QueueValidation(string key)
    {
        if (validationTimers.Remove(key, out var previous)) previous.Cancel();
        var timer = new CancellationTokenSource(); validationTimers[key] = timer;
        try
        {
            await Task.Delay(220, timer.Token);
            await liveEdits.WaitAsync(timer.Token);
            try
            {
                if (timer.IsCancellationRequested || session.GetPending(key) is null) return;
                var before = session.Snapshot.Object("draft").Copy();
                try
                {
                    await session.AcceptPendingAsync(key);
                    if (!JsonNode.DeepEquals(before.Object("config"), session.Snapshot.Object("draft").Object("config"))) editHistory.Push(before);
                    await session.ValidateAsync();
                }
                catch (Exception) { /* Field feedback owns validation errors; text stays editable. */ }
            }
            finally { liveEdits.Release(); }
            foreach (var update in fieldFeedback.Values.ToArray()) update();
            UpdateStatus();
        }
        catch (OperationCanceledException) { }
        finally { if (validationTimers.GetValueOrDefault(key) == timer) validationTimers.Remove(key); timer.Dispose(); }
    }

    private void BindFeedback(string key, string id, TextBlock feedback, Control input)
    {
        void Update()
        {
            var messages = session.Validation.Array("issues").OfType<JsonObject>()
                .Where(i => i.Text("severity") == "error" && i.Text("profile") == session.Profile &&
                    (i.Text("id") == id || i.Text("id").Replace("ThumbnailStartLocation", "ThumbnailMinimumSize") == id))
                .Select(i => i.Text("message"));
            var message = session.GetPending(key)?.Error ?? string.Join("\n", messages);
            if (session.Snapshot.Object("values")[id] is JsonObject state) input.IsEnabled = !state.Bool("inherited") && state.Bool("enabled", true);
            feedback.Text = message; feedback.IsVisible = message.Length != 0;
            feedback.Classes.Set("error", message.Length != 0);
            ToolTip.SetTip(input, message.Length != 0 ? message : null);
        }
        fieldFeedback[key] = Update; Update();
    }

    private async Task UndoLastAsync()
    {
        foreach (var timer in validationTimers.Values) timer.Cancel();
        await liveEdits.WaitAsync();
        try
        {
            if (session.Pending.LastOrDefault() is { } pending) await session.DiscardPendingAsync(pending.Key);
            else if (editHistory.TryPop(out var before)) await session.MutateAsync("restore-draft", new() { ["draft"] = before });
            await session.ProtectPendingAsync();
            await session.ValidateAsync();
        }
        finally { liveEdits.Release(); }
    }
    private async Task UndoCategoryAsync()
    {
        foreach (var edit in session.Pending.Where(p => p.Key.StartsWith(EditKey("")) || p.Key.StartsWith("global|" + sectionId + "|")).ToArray())
        {
            if (validationTimers.TryGetValue(edit.Key, out var timer)) timer.Cancel();
            await session.DiscardPendingAsync(edit.Key);
        }
        await liveEdits.WaitAsync();
        try {
            var before = session.Snapshot.Object("draft").Copy();
            await session.MutateAsync("undo-category", new() { ["section"] = sectionId, ["baseline"] = profileBaselines.GetValueOrDefault(session.Profile)?.DeepClone() });
            if (!JsonNode.DeepEquals(before.Object("config"), session.Snapshot.Object("draft").Object("config"))) editHistory.Push(before);
            await session.ValidateAsync();
        }
        finally { liveEdits.Release(); }
    }
    private TextBlock SearchText(string text)
    {
        var block = new TextBlock { TextWrapping = TextWrapping.Wrap };
        var words = (search.Text ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var offset = 0;
        while (offset < text.Length)
        {
            var match = words.Select(w => (word: w, index: text.IndexOf(w, offset, StringComparison.OrdinalIgnoreCase)))
                .Where(m => m.index >= 0).OrderBy(m => m.index).FirstOrDefault();
            if (match.word is null) { block.Inlines!.Add(new Run(text[offset..])); break; }
            if (match.index > offset) block.Inlines!.Add(new Run(text[offset..match.index]));
            block.Inlines!.Add(new Run(text.Substring(match.index, match.word.Length)) { Background = Brushes.Goldenrod, Foreground = Brushes.Black });
            offset = match.index + match.word.Length;
        }
        return block;
    }
}
