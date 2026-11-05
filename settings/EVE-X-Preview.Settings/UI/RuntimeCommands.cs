using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using EveX.Settings.Bridge;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    private async Task PresentCommandResultAsync(JsonObject command, JsonObject result)
    {
        switch (command.Text("id"))
        {
            case "running-windows": await PickWindowsAsync(result.Array("rows")); return;
            case "detected-sources":
                await Dialogs.MessageAsync(this, "Detected log sources", "Game logs\n" + result.Text("game", "Not found") + "\n\nChat logs\n" + result.Text("chat", "Not found")); return;
            case "shortcut-conflicts": await PresentConflictsAsync(result); return;
            case "update-status":
                SetAboutStatus(result);return;
            case "check-updates":
                activity = "Checking release information";
                for (var i = 0; i < 50; i++)
                {
                    await Task.Delay(200);
                    var status = await session.ReadAsync("command", new() { ["name"] = "update-status", ["arguments"] = new JsonObject() });
                    SetAboutStatus(status);
                    if (!status.Bool("busy")) {activity=status.Text("status");return;}
                }
                activity = "Update check is still running. Reopen About to inspect its status."; return;
        }
        activity = result.Text("message", "Action completed.");
        UpdateStatus();
    }
    private async Task PickWindowsAsync(JsonArray rows)
    {
        if (rows.Count == 0) { await Dialogs.MessageAsync(this, "Running applications", "No eligible running application windows were found."); return; }
        var content = new StackPanel { Spacing = 10 };
        content.Children.Add(Muted("Choose windows to add to an application group. Exact executable names and current titles remain editable in the members table."));
        var selected = new HashSet<int>();
        for (var i = 0; i < rows.Count; i++)
        {
            var index = i; var row = rows[i];
            var checkbox = new CheckBox { Content = new TextBlock { Text = row.Text("exe") + "\n" + row.Text("title"), TextWrapping = TextWrapping.Wrap } };
            checkbox.IsCheckedChanged += (_, _) => { if (checkbox.IsChecked == true) selected.Add(index); else selected.Remove(index); }; content.Children.Add(checkbox);
        }
        if (await Dialogs.ChoiceAsync(this, "Pick running application windows", content, [("Add to group", "add"), ("Cancel", "cancel")]) != "add" || selected.Count == 0) return;
        var groupKey = session.Profile + "|group|apps";
        var group = selections.GetValueOrDefault(groupKey, "");
        if (group == "")
        {
            if (await Dialogs.NameAsync(this, "Create application group") is not { } name) return;
            await session.MutateAsync("group", new() { ["kind"] = "apps", ["action"] = "add", ["name"] = name }); group = name; selections[groupKey] = group;
        }
        var members = await session.ReadAsync("rows", new() { ["kind"] = "app-group-members", ["group"] = group });
        var key = EditKey("rows|app-group-members|" + group);
        var additions = new JsonArray();
        foreach (var index in selected) additions.Add(new JsonArray(rows[index].Text("exe"), rows[index].Text("title")));
        session.AppendPendingRows(key, group + " members", "app-group-members", group, members.Array("rows"), additions);
        await session.ProtectPendingAsync();
        QueueValidation(key); activity = "Selected application windows added to the group.";
    }
    private async Task PresentConflictsAsync(JsonObject result)
    {
        var groups = result.Array("groups");
        var lines = new List<string>();
        foreach (var group in groups)
        {
            if (group is JsonObject entry)
            {
                lines.Add(string.Join(" / ", entry.Array("keys").Select(JsonExtensions.Display)) is { Length: > 0 } keys ? keys : "Shortcut conflict");
                foreach (var binding in entry.Array("entries")) lines.Add("  " + binding.Text("label") + " · " + binding.Text("key"));
            }
            else if (group is JsonArray bindings)
                foreach (var binding in bindings) lines.Add(binding.Text("label") + " · " + binding.Text("key"));
        }
        await Dialogs.MessageAsync(this, "Shortcut conflicts", lines.Count == 0 ? "No shortcut conflicts were reported by the application." : string.Join("\n", lines));
    }
}
