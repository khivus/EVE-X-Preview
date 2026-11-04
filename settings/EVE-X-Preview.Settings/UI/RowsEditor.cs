using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Input.Platform;
using EveX.Settings.Bridge;
using EveX.Settings.State;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    private readonly Dictionary<string, Stack<JsonArray>> rowUndo = new();
    private async void LoadRowsEditor(StackPanel host, string kind, string group, bool inherited, int generation)
    {
        try
        {
            host.Children.Add(Muted("Loading rows…"));
            var data = await session.ReadAsync("rows", new() { ["kind"] = kind, ["group"] = group });
            if (generation != pageGeneration) return;
            BuildRowsEditor(host, kind, group, data, inherited);
        }
        catch (Exception e) { if (generation == pageGeneration) { host.Children.Clear(); host.Children.Add(new TextBlock { Text = e.Message, Classes = { "error" } }); } }
    }
    private void BuildRowsEditor(StackPanel host, string kind, string group, JsonObject data, bool inherited)
    {
        host.Children.Clear();
        var columns = data.Array("columns").OfType<JsonObject>().ToArray();
        var key = EditKey("rows|" + kind + "|" + group);
        var rows = (JsonArray)(session.GetPending(key)?.Arguments["rows"] as JsonArray ?? data.Array("rows")).DeepClone();
        var readOnlyRows = columns.Any(column => column.Text("kind") == "readonly");
        if (!rowUndo.ContainsKey(key)) rowUndo[key] = new();
        var selected = Math.Clamp(int.TryParse(selections.GetValueOrDefault(key), out var index) ? index : 0, 0, Math.Max(0, rows.Count - 1));
        var list = new StackPanel { Spacing = 7, MinWidth = Math.Max(480, columns.Length * 190 + 46) };
        var description = new TextBlock { Text = "Select a row number to reorder or remove it. Copy and paste uses tab-separated columns. Accepted row changes participate in the shared Undo history.", Classes = { "muted" } }; host.Children.Add(description);
        var toolbar = new WrapPanel(); host.Children.Add(toolbar);
        void StageRows()
        {
            Stage(key, group == "" ? kind : group + " members", "rows", new() { ["kind"] = kind, ["group"] = group, ["rows"] = rows.DeepClone() });
        }
        void Remember() { rowUndo[key].Push((JsonArray)rows.DeepClone()); while (rowUndo[key].Count > 50) { var kept = rowUndo[key].Take(50).Reverse().ToArray(); rowUndo[key].Clear(); foreach (var entry in kept) rowUndo[key].Push(entry); } }
        Button Tool(string title, Func<Task> action, bool enabled = true)
        {
            var button = ActionButton(title, action, false); button.Margin = new Thickness(0, 0, 8, 8); button.IsEnabled = enabled && !inherited && !data.Bool("inherited"); toolbar.Children.Add(button); return button;
        }
        void RenderRows()
        {
            list.Children.Clear();
            var header = new Grid { ColumnDefinitions = new ColumnDefinitions("42," + string.Join(',', columns.Select(_ => "*"))) };
            for (var c = 0; c < columns.Length; c++) { var title = new TextBlock { Text = columns[c].Text("label"), FontWeight = FontWeight.SemiBold, Margin = new Thickness(5), TextWrapping = TextWrapping.Wrap }; Grid.SetColumn(title, c + 1); header.Children.Add(title); }
            list.Children.Add(header);
            for (var r = 0; r < rows.Count; r++)
            {
                var rowIndex = r; var values = rows[r] as JsonArray ?? new JsonArray();
                var line = new Grid { ColumnDefinitions = new ColumnDefinitions("42," + string.Join(',', columns.Select(_ => "*"))) };
                var select = new Button { Content = (r + 1).ToString(), Padding = new Thickness(4), HorizontalAlignment = HorizontalAlignment.Stretch, Background = selected == r ? new SolidColorBrush(Color.Parse("#3265A0")) : Brushes.Transparent };
                select.Click += (_, _) => { selected = rowIndex; selections[key] = selected.ToString(); RenderRows(); }; line.Children.Add(select);
                for (var c = 0; c < columns.Length; c++)
                {
                    var columnIndex = c;
                    var control = TypedInput(columns[c], c < values.Count ? values[c] : null, value =>
                    {
                        Remember();
                        while (values.Count <= columnIndex) values.Add("");
                        values[columnIndex] = value?.DeepClone(); StageRows();
                    });
                    control.Margin = new Thickness(5, 0); control.IsEnabled = !inherited && !data.Bool("inherited"); Grid.SetColumn(control, c + 1); line.Children.Add(control);
                }
                list.Children.Add(line);
            }
            if (rows.Count == 0) list.Children.Add(Muted("No rows yet. Add one or paste rows from the clipboard."));
        }
        Tool("Add row", () =>
        {
            Remember(); var values = new JsonArray(); foreach (var column in columns) values.Add(column["default"]?.DeepClone() ?? (column.Text("kind") == "bool" ? JsonValue.Create(false) : JsonValue.Create(""))); rows.Add(values); selected = rows.Count - 1; StageRows(); RenderRows(); return Task.CompletedTask;
        }, !readOnlyRows);
        Tool("Remove row", () => { if (rows.Count > 0) { Remember(); rows.RemoveAt(Math.Min(selected, rows.Count - 1)); selected = Math.Max(0, selected - 1); StageRows(); RenderRows(); } return Task.CompletedTask; }, !readOnlyRows);
        Tool("Move up", () => { if (selected > 0 && selected < rows.Count) { Remember(); var row = rows[selected]!.DeepClone(); rows.RemoveAt(selected); rows.Insert(--selected, row); StageRows(); RenderRows(); } return Task.CompletedTask; }, !readOnlyRows);
        Tool("Move down", () => { if (selected < rows.Count - 1) { Remember(); var row = rows[selected]!.DeepClone(); rows.RemoveAt(selected); rows.Insert(++selected, row); StageRows(); RenderRows(); } return Task.CompletedTask; }, !readOnlyRows);
        Tool("Copy all", async () => { if (Clipboard is { } clipboard) await clipboard.SetTextAsync(TableClipboard.Write(rows)); });
        Tool("Copy row", async () => { if (Clipboard is { } clipboard && rows.Count != 0) await clipboard.SetTextAsync(TableClipboard.Write(new JsonArray(rows[Math.Min(selected, rows.Count - 1)]!.DeepClone()))); });
        Tool("Paste rows", async () => { if (Clipboard is { } clipboard && await clipboard.TryGetTextAsync() is { } text) { var pasted = TableClipboard.Read(text, columns.Length); Remember(); foreach (var row in pasted) rows.Add(row?.DeepClone()); StageRows(); RenderRows(); } }, !readOnlyRows);
        if (columns.FirstOrDefault()?.Text("kind") is "character" or "exe")
        {
            var applications = columns[0].Text("kind") == "exe";
            Tool(applications ? "Add running applications…" : "Add open EVE characters…", async () =>
            {
                var windows = await session.ReadAsync("command", new() { ["name"] = "running-windows", ["arguments"] = new JsonObject { ["applications"] = applications } });
                var selectedWindows = await ChooseRunningRowsAsync(windows.Array("rows"), applications);
                if (selectedWindows.Count == 0) return;
                Remember();
                foreach (var window in selectedWindows)
                {
                    var values = new JsonArray();
                    for (var c = 0; c < columns.Length; c++)
                        values.Add(c == 0 ? JsonValue.Create(window.Text(applications ? "exe" : "character")) : applications && c == 1 ? JsonValue.Create(window.Text("title")) : columns[c]["default"]?.DeepClone() ?? (columns[c].Text("kind") == "bool" ? JsonValue.Create(false) : JsonValue.Create("")));
                    rows.Add(values);
                }
                StageRows(); RenderRows();
            }, session.Role != "offline");
        }
        Tool("Undo row edit", () => { if (rowUndo[key].TryPop(out var previous)) { rows = (JsonArray)previous.DeepClone(); StageRows(); RenderRows(); } return Task.CompletedTask; });
        host.Children.Add(new ScrollViewer { Content = list, HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto, VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled });
        var save = ActionButton("Accept rows to draft", async () => { await session.AcceptPendingAsync(key); rowUndo.Remove(key); }); save.IsEnabled = !inherited && !data.Bool("inherited"); host.Children.Add(save);
        if (session.GetPending(key) is not null) host.Children.Add(ActionButton("Discard unfinished row edits", () => session.DiscardPendingAsync(key)));
        if (session.GetPending(key)?.Error is { } error) host.Children.Add(new TextBlock { Text = error, Classes = { "error" } });
        RenderRows();
    }
    private async Task<List<JsonNode>> ChooseRunningRowsAsync(JsonArray rows, bool applications)
    {
        if (rows.Count == 0) { await Dialogs.MessageAsync(this, "Running windows", "No matching windows were found. You can enter names manually in the table."); return []; }
        var selected = new HashSet<int>(); var content = new StackPanel { Spacing = 10 };
        content.Children.Add(Muted("Select windows to append. Existing assignments and unfinished edits are kept."));
        for (var i = 0; i < rows.Count; i++)
        {
            var index = i; var item = rows[i];
            var check = new CheckBox { Content = item.Text(applications ? "exe" : "character") + " · " + item.Text("title") };
            check.IsCheckedChanged += (_, _) => { if (check.IsChecked == true) selected.Add(index); else selected.Remove(index); }; content.Children.Add(check);
        }
        return await Dialogs.ChoiceAsync(this, applications ? "Choose running applications" : "Choose open EVE characters", content, [("Append selected", "append"), ("Cancel", "cancel")]) == "append" ? selected.Select(index => rows[index]!).ToList() : [];
    }
}
