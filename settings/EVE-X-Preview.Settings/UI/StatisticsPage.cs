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
    private readonly Dictionary<string, int> statisticsListLimits = new();
    private readonly Dictionary<string, Vector> statisticsListOffsets = new();

    private void RenderStatistics()
    {
        var dashboardHost = new ContentControl();
        var status = Muted("");
        var start = new Button { Content = T("Scan all logs"), IsEnabled = false };
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
                if (view.Bool("extendedNeedsRescan")) status.Text += "\n" + T("Run an archive scan to calculate logistics, fleet, burst, module and system statistics for older logs.");
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
                var content = stats.Array("categories").ToJsonString() + view.Array("materials").ToJsonString() + view.Array("bursts").ToJsonString() + view.Array("drones").ToJsonString() + view.Array("modules").ToJsonString() + view.Array("systems").ToJsonString();
                if (content == lastContent) return;
                lastContent = content;
                // Build off-tree so polling never exposes an empty dashboard to layout.
                var dashboard = new StackPanel { Spacing = 12 };
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
                            tile.Children.Add(Muted(leader));
                        else if (requestedCharacter == "" && category.Text("title") != "Saved configuration" && metric.Number("value") > 0)
                            tile.Children.Add(Muted("No character history available"));
                        if (metric.Text("counterpart") is { Length: > 0 } counterpart)
                            tile.Children.Add(new TextBlock { Text = counterpart, MaxWidth = 360,
                                TextWrapping = TextWrapping.Wrap, Classes = { "muted" } });
                        metrics.Children.Add(tile);
                    }
                    group.Children.Add(metrics);
                    if (category.Text("title") == "Command and mining bursts" && view.Array("bursts")
                        .OrderByDescending(e => e.Number("count")).ThenBy(e => e.Text("name")).FirstOrDefault() is { } burst)
                    {
                        var tile = new StackPanel { Spacing = 4, Margin = new Thickness(0, 0, 24, 6), MaxWidth = 280 };
                        tile.Children.Add(new TextBlock { Text = burst.Text("name"), FontSize = 22,
                            FontWeight = FontWeight.SemiBold, TextWrapping = TextWrapping.Wrap });
                        tile.Children.Add(Muted("Most used burst type"));
                        tile.Children.Add(Muted(burst.Text("count") + " " + T("applications")));
                        if (requestedCharacter == "") tile.Children.Add(Muted(burst.Text("character")));
                        metrics.Children.Add(tile);
                    }
                    dashboard.Children.Add(Card(group));
                }
                foreach (var (key, title, label) in new[] {
                    ("modules", "Most used modules", "Module"),
                    ("drones", "Most used drones / fighters", "Drone / fighter"),
                    ("systems", "Most visited systems", "System") })
                {
                    var entries = view.Array(key).OrderByDescending(e => e.Number("count")).ThenBy(e => e.Text("name")).ToArray();
                    if (entries.Length == 0) continue;
                    var table = StatisticsRanking(key, title, "2*,*,2*", [label, "Records", "Character"],
                        entries.Select(e => new[] { e.Text("name"), e.Text("count"), requestedCharacter == "" ? e.Text("character") : "" }).ToArray());
                    dashboard.Children.Add(Card(table));
                }
                if (view.Array("materials") is { Count: > 0 } materials)
                {
                    var entries = materials.OrderByDescending(m => m.Number("units")).ThenBy(m => m.Text("name"));
                    dashboard.Children.Add(Card(StatisticsRanking("materials", "Mined materials", "2*,*,*,2*",
                        ["Material", "Mined units", "Residue units", "Character"],
                        entries.Select(m => new[] { m.Text("name"), m.Text("units"), m.Text("residue"),
                            requestedCharacter == "" ? m.Text("character") : "" }).ToArray())));
                }
                var offset = pageScroll.Offset;
                dashboardHost.Content = dashboard;
                // Recalculate the extent before restoring; a shorter result is clamped normally.
                pageScroll.UpdateLayout();
                pageScroll.Offset = offset;
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
        dashboardHost.DetachedFromVisualTree += (_, _) => timer.Stop();
        page.Children.Add(filters);
        page.Children.Add(Muted("Character and date filters apply to EVE log statistics."));
        page.Children.Add(Muted("With All characters selected, each character name identifies the leader for that statistic."));
        page.Children.Add(actions);
        page.Children.Add(Muted("Rescans are faster: unchanged logs are skipped and changed logs resume from saved progress."));
        page.Children.Add(status);
        page.Children.Add(progress);
        page.Children.Add(dashboardHost);
        _ = RefreshAsync();
    }

    private Control StatisticsRanking(string key, string title, string columns, string[] headings, string[][] entries)
    {
        // Only rows scroll. The title, amount selector and column labels stay visible.
        var table = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto"), ClipToBounds = true };
        var headerContent = new StackPanel { Spacing = 8 };
        var heading = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto") };
        heading.Children.Add(new TextBlock { Text = T(title), FontWeight = FontWeight.SemiBold,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center });
        var show = Muted("Show");
        show.VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center;
        show.Margin = new Thickness(12, 0, 8, 0);
        Grid.SetColumn(show, 1); heading.Children.Add(show);
        int[] choices = [5, 10, 25, 50, 100, int.MaxValue];
        var limit = statisticsListLimits.GetValueOrDefault(key, 5);
        var amount = new ComboBox { MinWidth = 85, ItemsSource = choices.Select(n => n == int.MaxValue ? T("All") : n.ToString()).ToArray(),
            SelectedIndex = Array.IndexOf(choices, limit) };
        Grid.SetColumn(amount, 2); heading.Children.Add(amount);
        headerContent.Children.Add(heading);
        TableHeading(headerContent, columns, headings);
        var header = new Border { Child = headerContent, Padding = new Thickness(0, 8),
            Classes = { "statistics-header" } };
        var rows = new StackPanel();
        var viewport = new ScrollViewer { Content = rows,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled };
        Grid.SetRow(viewport, 1);
        table.Children.Add(viewport); table.Children.Add(header);

        void Populate(bool resetOffset = true)
        {
            var offset = resetOffset ? default : statisticsListOffsets.GetValueOrDefault(key);
            if (resetOffset) statisticsListOffsets[key] = default;
            rows = new StackPanel();
            viewport.Content = rows;
            var count = Math.Min(entries.Length, statisticsListLimits.GetValueOrDefault(key, 5));
            viewport.VerticalScrollBarVisibility = count > 10 ? Avalonia.Controls.Primitives.ScrollBarVisibility.Auto
                : Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled;
            viewport.IsScrollChainingEnabled = count <= 10;
            viewport.MaxHeight = count > 10 ? 360 : double.PositiveInfinity;
            viewport.Offset = default;
            foreach (var values in entries.Take(statisticsListLimits.GetValueOrDefault(key, 5)))
            {
                var row = TableLine(columns);
                for (var column = 0; column < values.Length; column++)
                    row.Children.Add(TableCell(Muted(values[column]), column));
                rows.Children.Add(row);
            }
            if (!resetOffset) Dispatcher.UIThread.Post(() => viewport.Offset = offset, DispatcherPriority.Loaded);
        }
        var attached = false;
        viewport.AttachedToVisualTree += (_, _) => attached = true;
        viewport.DetachedFromVisualTree += (_, _) => attached = false;
        viewport.ScrollChanged += (_, _) => { if (attached) statisticsListOffsets[key] = viewport.Offset; };
        viewport.LayoutUpdated += (_, _) =>
        {
            headerContent.Margin = new Thickness(0, 0, Math.Max(0, viewport.Bounds.Width - viewport.Viewport.Width), 0);
            if (rows.Children.Count <= 10 || rows.Children[9].Bounds.Height <= 0) return;
            // Account for wrapped names and varying font sizes rather than assuming
            // all rows have the same height. Fit exactly the first ten rows.
            var height = rows.Children[9].Bounds.Bottom + rows.Children[9].Margin.Bottom;
            if (Math.Abs(viewport.MaxHeight - height) > 0.5) viewport.MaxHeight = height;
        };
        amount.SelectionChanged += (_, _) =>
        {
            if (amount.SelectedIndex < 0) return;
            statisticsListLimits[key] = choices[amount.SelectedIndex];
            Populate();
        };
        Populate(false);
        return table;
    }

}
