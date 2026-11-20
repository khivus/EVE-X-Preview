using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using EveX.Settings.Bridge;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    private string statisticsCharacter = "";
    private string statisticsRange = "all";
    private string[] statisticsCharacters = [];

    private void RenderStatistics()
    {
        var dashboard = new StackPanel { Spacing = 12 };
        var status = Muted("");
        var start = new Button { Content = T("Scan or resume log archives"), IsEnabled = false };
        var stop = new Button { Content = T("Stop archive scan"), IsEnabled = false };
        var refresh = new Button { Content = T("Refresh statistics") };
        var actions = new WrapPanel { Children = { start, stop, refresh } };
        foreach (var button in actions.Children)
            button.Margin = new Thickness(0, 0, 8, 8);
        var ranges = new (string Key, string Label)[] {
            ("all", "All time"), ("session", "Last session"), ("day", "Today (UTC)"),
            ("week", "Last 7 days"), ("month", "Last 30 days"), ("year", "Last 365 days")
        };
        var character = new ComboBox { Width = 240, MaxWidth = 240 };
        var date = new ComboBox { MinWidth = 145, ItemsSource = ranges.Select(r => T(r.Label)).ToArray(),
            SelectedIndex = Array.FindIndex(ranges, r => r.Key == statisticsRange) };
        var filters = new WrapPanel();
        filters.Children.Add(new StackPanel { Margin = new Thickness(0, 0, 14, 8), Spacing = 4, Children = { Muted("Character"), character } });
        filters.Children.Add(new StackPanel { Margin = new Thickness(0, 0, 14, 8), Spacing = 4, Children = { Muted("Date range"), date } });
        var progress = new ProgressBar { Minimum = 0, Maximum = 100, IsVisible = false };
        var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
        var viewGeneration = pageGeneration;
        var readGate = new SemaphoreSlim(1, 1);
        bool settingCharacters = false;
        string lastContent = "";

        void ShowCharacters()
        {
            settingCharacters = true;
            character.ItemsSource = new[] { T("All characters") }.Concat(statisticsCharacters).ToArray();
            character.SelectedIndex = Array.IndexOf(statisticsCharacters, statisticsCharacter) + 1;
            settingCharacters = false;
        }
        ShowCharacters();

        async Task RefreshAsync(string command = "session-statistics", bool force = false)
        {
            if (command == "session-statistics" && !force)
            {
                if (!readGate.Wait(0)) return;
            }
            else await readGate.WaitAsync();
            if (viewGeneration != pageGeneration)
            {
                readGate.Release();
                return;
            }
            if (force) refresh.IsEnabled = false;
            var requestedCharacter = statisticsCharacter;
            var requestedRange = statisticsRange;
            try
            {
                var stats = await session.ReadAsync("command", new() { ["name"] = command, ["arguments"] = new JsonObject
                    { ["character"] = requestedCharacter, ["range"] = requestedRange, ["refresh"] = force } });
                if (viewGeneration != pageGeneration) return;
                var available = stats.Bool("available", stats.Bool("enabled"));
                var scanning = stats.Bool("scanning");
                start.IsEnabled = available && !scanning;
                stop.IsEnabled = available;
                filters.IsEnabled = available;
                var view = stats.Object("view");
                status.Text = T(stats.Text("scanStatus"));
                if (!stats.Bool("enabled"))
                    status.Text = (T("Live statistics collection is off.") + "\n" + status.Text).TrimEnd();
                var progressText = stats.Text("progressText");
                if (progressText != "") status.Text += "\n" + progressText;
                if (view.Bool("pending")) status.Text += "\n" + T("Reading the statistics index…");
                if (view.Bool("deferred")) status.Text += "\n" + T("Filtered results update when the scan finishes or stops.");
                if (view.Text("error") is { Length: > 0 } viewError) status.Text += "\n" + T(viewError);
                if (view.Bool("combatNeedsRescan")) status.Text += "\n" + T("Run an archive scan to calculate combat peaks for older indexed logs.");
                if (view.Bool("extendedNeedsRescan")) status.Text += "\n" + T("Run an archive scan to calculate logistics, fleet, module and system statistics for older logs.");
                var scan = stats.Object("progress");
                var total = scan.Int("total");
                progress.IsVisible = scanning || total > 0;
                progress.IsIndeterminate = scanning && total == 0;
                progress.Value = total > 0 ? 100.0 * scan.Int("scanned") / total : 0;
                if (available)
                {
                    timer.Interval = TimeSpan.FromMilliseconds(view.Bool("pending") ? 350 : 2000);
                    timer.Start();
                }
                if (requestedCharacter != statisticsCharacter || requestedRange != statisticsRange) return;
                if (view["characters"] is JsonArray names)
                {
                    var indexed = names.Select(JsonExtensions.Display).Order(StringComparer.CurrentCultureIgnoreCase).ToArray();
                    if (!indexed.SequenceEqual(statisticsCharacters)) { statisticsCharacters = indexed; ShowCharacters(); }
                }
                var content = stats.Array("categories").ToJsonString() + view.Array("materials").ToJsonString() + view.Array("drones").ToJsonString() + view.Array("modules").ToJsonString() + view.Array("systems").ToJsonString();
                if (content == lastContent) return;
                lastContent = content;
                dashboard.Children.Clear();
                foreach (var category in stats.Array("categories"))
                {
                    var group = new StackPanel { Spacing = 12 };
                    group.Children.Add(new TextBlock { Text = T(category.Text("title")), FontWeight = FontWeight.SemiBold });
                    var metrics = new WrapPanel();
                    foreach (var metric in category.Array("metrics"))
                    {
                        var tile = new StackPanel { Spacing = 4, Margin = new Thickness(0, 0, 24, 6) };
                        tile.Children.Add(new TextBlock { Text = metric.Text("value"), FontSize = 30, FontWeight = FontWeight.SemiBold });
                        tile.Children.Add(Muted(metric.Text("label")));
                        if (requestedCharacter == "" && metric.Text("character") is { Length: > 0 } leader)
                            tile.Children.Add(Muted(T("Top character") + ": " + leader));
                        else if (requestedCharacter == "" && category.Text("title") != "Saved configuration" && metric.Number("value") > 0)
                            tile.Children.Add(Muted("No character history available"));
                        metrics.Children.Add(tile);
                    }
                    group.Children.Add(metrics);
                    dashboard.Children.Add(Card(group));
                }
                foreach (var (key, title, label, limit) in new[] {
                    ("modules", "Most used module", "Module", 1),
                    ("drones", "Most used drone / fighter", "Drone / fighter", 1),
                    ("systems", "Most visited systems (top 10)", "System", 10) })
                {
                    var entries = view.Array(key).OrderByDescending(e => e.Number("count")).ThenBy(e => e.Text("name")).Take(limit).ToArray();
                    if (entries.Length == 0) continue;
                    var group = new StackPanel { Spacing = 8 };
                    group.Children.Add(new TextBlock { Text = T(title), FontWeight = FontWeight.SemiBold });
                    TableHeading(group, "2*,*,2*", label, "Records", "Top character");
                    foreach (var entry in entries)
                    {
                        var row = TableLine("2*,*,2*");
                        row.Children.Add(TableCell(Muted(entry.Text("name")), 0));
                        row.Children.Add(TableCell(Muted(entry.Text("count")), 1));
                        row.Children.Add(TableCell(Muted(requestedCharacter == "" ? entry.Text("character") : ""), 2));
                        group.Children.Add(row);
                    }
                    dashboard.Children.Add(Card(group));
                }
                if (view.Array("materials") is { Count: > 0 } materials)
                {
                    var group = new StackPanel { Spacing = 8 };
                    group.Children.Add(new TextBlock { Text = T("Mined materials"), FontWeight = FontWeight.SemiBold });
                    TableHeading(group, "2*,*,*,2*", "Material", "Mined units", "Residue units", "Top character");
                    var clipboard = new TableCellClipboard(this, group);
                    foreach (var material in materials.OrderByDescending(m => m.Number("units")))
                    {
                        var row = TableLine("2*,*,*,2*");
                        row.Children.Add(TableCell(Muted(material.Text("name")), 0));
                        row.Children.Add(TableCell(Muted(material.Text("units")), 1));
                        row.Children.Add(TableCell(Muted(material.Text("residue")), 2));
                        row.Children.Add(TableCell(Muted(requestedCharacter == "" ? material.Text("character") : ""), 3));
                        group.Children.Add(row);
                        clipboard.Add(row, 0, () => material.Text("name"));
                        clipboard.Add(row, 1, () => material.Text("units"));
                        clipboard.Add(row, 2, () => material.Text("residue"));
                    }
                    dashboard.Children.Add(Card(group));
                }
            }
            catch (Exception error)
            {
                status.Text = T(error.Message);
            }
            finally
            {
                if (force) refresh.IsEnabled = true;
                readGate.Release();
                if (viewGeneration == pageGeneration &&
                    (requestedCharacter != statisticsCharacter || requestedRange != statisticsRange))
                    Dispatcher.UIThread.Post(() => _ = RefreshAsync());
            }
        }
        start.Click += async (_, _) => await RefreshAsync("statistics-scan-start");
        stop.Click += async (_, _) => await RefreshAsync("statistics-scan-stop");
        refresh.Click += async (_, _) => await RefreshAsync(force: true);
        character.SelectionChanged += async (_, _) =>
        {
            if (settingCharacters || character.SelectedIndex < 0) return;
            statisticsCharacter = character.SelectedIndex > 0 ? statisticsCharacters[character.SelectedIndex - 1] : "";
            await RefreshAsync();
        };
        date.SelectionChanged += async (_, _) =>
        {
            if (date.SelectedIndex < 0) return;
            statisticsRange = ranges[date.SelectedIndex].Key;
            await RefreshAsync();
        };
        timer.Tick += async (_, _) =>
        {
            if (viewGeneration != pageGeneration || closing) { timer.Stop(); return; }
            await RefreshAsync();
        };
        dashboard.DetachedFromVisualTree += (_, _) => timer.Stop();
        page.Children.Add(filters);
        page.Children.Add(Muted("Character and date filters apply to EVE log statistics."));
        page.Children.Add(actions);
        page.Children.Add(status);
        page.Children.Add(progress);
        page.Children.Add(dashboard);
        _ = RefreshAsync();
    }
}
