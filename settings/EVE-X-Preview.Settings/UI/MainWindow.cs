using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using EveX.Settings.Bridge;
using EveX.Settings.State;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow : Window
{
    private readonly EditorSession session;
    private readonly SingleInstance instance;
    private readonly StackPanel navigation = new() { Spacing = 3, Margin = new Thickness(16, 6) };
    private readonly StackPanel page = new() { Spacing = 16, Margin = new Thickness(28, 22, 28, 30) };
    private readonly ScrollViewer pageScroll = new() { HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled };
    private readonly ComboBox profiles = new() { MinWidth = 175, MaxWidth = 320, HorizontalAlignment = HorizontalAlignment.Stretch };
    private readonly TextBox search = new() { PlaceholderText = "Search settings…", Margin = new Thickness(16, 12), InnerLeftContent = new TextBlock { Text = "⌕", FontSize = 22, Margin = new Thickness(10, 0) } };
    private readonly TextBlock connection = Muted("Connecting to canonical settings…");
    private readonly TextBlock draft = new() { FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock error = new() { Classes = { "error" }, IsVisible = false, Margin = new Thickness(28, 8) };
    private readonly TextBlock notice = Muted("");
    private readonly Button apply, discard, accept, undo, reconnect;
    private readonly Dictionary<string, Vector> scrollPositions = new();
    private readonly Dictionary<string, string> selections = new();
    private string sectionId = "appearance";
    private bool busy, closing, closePromptOpen, updatingProfiles, rendering;
    private string activity = "";
    private string? captureTheme;

    public MainWindow(string installation, SingleInstance instance)
    {
        this.instance = instance;
        session = new EditorSession(new BridgeTransport(installation));
        Width = 1180; Height = 820; MinWidth = 760; MinHeight = 560;
        Title = "EVE-X-Preview · Settings";
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,*,Auto,Auto") };
        var header = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), Margin = new Thickness(26, 20, 26, 18) };
        var branding = new StackPanel { Spacing = 5 };
        branding.Children.Add(new TextBlock { Text = "EVE-X-PREVIEW", FontSize = 13, LetterSpacing = 2.5, FontWeight = FontWeight.SemiBold });
        branding.Children.Add(new TextBlock { Text = "Make your workspace yours.", FontSize = 24, FontWeight = FontWeight.SemiBold });
        branding.Children.Add(connection);
        header.Children.Add(branding);
        var profilePanel = new StackPanel { Spacing = 6, Margin = new Thickness(20, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        profilePanel.Children.Add(Muted("EDITING PROFILE")); profilePanel.Children.Add(profiles);
        Grid.SetColumn(profilePanel, 1); header.Children.Add(profilePanel); root.Children.Add(header);
        var body = new Grid { ColumnDefinitions = new ColumnDefinitions("242,*") };
        var sidebar = new DockPanel(); DockPanel.SetDock(search, Dock.Top); sidebar.Children.Add(search); sidebar.Children.Add(new ScrollViewer { Content = navigation });
        body.Children.Add(new Border { Child = sidebar, Classes = { "sidebar" }, BorderThickness = new Thickness(0, 1, 1, 1), Opacity = 0.99 });
        pageScroll.Content = page; Grid.SetColumn(pageScroll, 1); body.Children.Add(pageScroll); Grid.SetRow(body, 1); root.Children.Add(body);
        Grid.SetRow(error, 2); root.Children.Add(error);
        var footer = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), Margin = new Thickness(26, 16) };
        var state = new StackPanel { Spacing = 5, VerticalAlignment = VerticalAlignment.Center }; state.Children.Add(draft); state.Children.Add(notice); footer.Children.Add(state);
        var actions = new WrapPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
        reconnect = ActionButton("Reconnect", () => ConnectAsync(), false);
        undo = ActionButton("Undo", async () => { await session.MutateAsync("undo", new()); }, true);
        accept = ActionButton("Accept edits", () => session.AcceptPendingAsync(), true);
        discard = ActionButton("Discard", DiscardAsync, true);
        apply = ActionButton("Apply & restart", ApplyAsync, false); apply.Classes.Add("accent");
        foreach (var button in new[] { reconnect, undo, accept, discard, apply }) { button.Margin = new Thickness(6, 0, 0, 0); actions.Children.Add(button); }
        Grid.SetColumn(actions, 1); footer.Children.Add(actions); Grid.SetRow(footer, 3); root.Children.Add(footer); Content = root;
        search.TextChanged += (_, _) => RenderNavigation();
        profiles.SelectionChanged += async (_, _) =>
        {
            if (updatingProfiles || profiles.SelectedItem is not string profile || profile == session.Profile) return;
            SaveScroll(); await RunAsync(async () => { await session.MutateAsync("select-profile", new() { ["name"] = profile }); });
        };
        session.Changed += () => Dispatcher.UIThread.Post(UpdateStatus);
        instance.ActivationRequested += ActivateRequested;
        Closing += OnClosing;
        Closed += async (_, _) => { instance.ActivationRequested -= ActivateRequested; await session.DisposeAsync(); };
        Opened += async (_, _) =>
        {
            await RunAsync(ConnectAsync);
            if (Program.CaptureDirectory is { } directory && session.Connected)
            {
                try { await CaptureFixturesAsync(directory); }
                catch (Exception error) { ShowError(error); }
            }
        };
        KeyDown += async (_, args) =>
        {
            if (args.KeyModifiers == KeyModifiers.Control && args.Key == Key.F) { search.Focus(); args.Handled = true; }
            if (args.KeyModifiers == KeyModifiers.Control && args.Key == Key.S && apply.IsEnabled) { args.Handled = true; await RunAsync(ApplyAsync); }
        };
        UpdateStatus();
    }
    private void ActivateRequested() => Dispatcher.UIThread.Post(() => { if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal; Show(); Activate(); });
    private async Task ConnectAsync()
    {
        await session.ConnectAsync();
        if (session.Metadata?.Sections.All(s => s.Text("id") != sectionId) == true) sectionId = session.Metadata.Sections.FirstOrDefault().Text("id");
        RenderNavigation(); RenderPage();
    }
    private Button ActionButton(string label, Func<Task> action, bool render = true)
    {
        var button = new Button { Content = label };
        button.Click += async (_, _) => await RunAsync(action, render);
        return button;
    }
    private async Task RunAsync(Func<Task> action, bool render = true)
    {
        if (busy || closing) return;
        busy = true; error.IsVisible = false; UpdateStatus();
        try { await action(); if (render) { RenderNavigation(); RenderPage(); } }
        catch (Exception e) { ShowError(e); }
        finally { busy = false; UpdateStatus(); }
    }
    private void ShowError(Exception exception)
    {
        error.Text = exception.Message; error.IsVisible = true;
        if (exception is BridgeException { Details: { } details })
        {
            var items = details as JsonArray ?? details.Array("issues");
            foreach (var item in items.Take(8)) error.Text += "\n" + (item is JsonObject ? item.Text("message") : JsonExtensions.Display(item));
        }
    }
    private void UpdateStatus()
    {
        connection.Text = session.Connected ? (session.Role == "offline" ? "Offline settings · canonical host · runtime actions unavailable" : "Connected to EVE-X-Preview · shared settings system") : "Disconnected · draft and unfinished edits retained";
        var pending = session.Pending.Count;
        draft.Text = pending != 0 ? $"{pending} unfinished edit{(pending == 1 ? "" : "s")}" : session.Dirty ? "Unapplied changes" : "All changes applied";
        notice.Text = busy ? "Working…" : pending != 0 ? "Accept edits to validate them with the application." : activity != "" ? activity : "Apply saves through the application and restarts it when needed.";
        apply.IsEnabled = !busy && session.Connected && session.Dirty && pending == 0;
        accept.IsEnabled = !busy && session.Connected && pending != 0;
        discard.IsEnabled = !busy && session.Connected && session.Dirty;
        undo.IsEnabled = !busy && session.Connected && session.Snapshot.Bool("canUndo") && pending == 0;
        reconnect.IsVisible = !session.Connected; reconnect.IsEnabled = !busy;
        profiles.IsEnabled = !busy && session.Connected;
        page.IsEnabled = !busy && session.Connected;
        if (!rendering)
        {
            updatingProfiles = true;
            var names = session.Snapshot.Array("profiles").Select(item => item is JsonObject ? item.Text("name") : JsonExtensions.Display(item)).ToArray();
            if (!names.SequenceEqual((profiles.ItemsSource as string[]) ?? [])) profiles.ItemsSource = names;
            profiles.SelectedItem = session.Profile; updatingProfiles = false;
        }
        if (session.Snapshot.Object("values")["SettingsTheme"] is { } theme)
        {
            var requested = session.Pending.FirstOrDefault(edit => edit.Arguments.Text("id") == "SettingsTheme")?.Arguments.Text("value") ?? theme.Text("value");
            ThemeController.Apply(captureTheme ?? requested, session.Metadata);
        }
    }
    private string ScrollKey => session.Profile + "|" + sectionId;
    private void SaveScroll() => scrollPositions[ScrollKey] = pageScroll.Offset;
    private void RenderNavigation()
    {
        navigation.Children.Clear();
        if (session.Metadata is not { } metadata) return;
        var query = (search.Text ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var group in metadata.Sections.GroupBy(s => s.Text("page")))
        {
            var visible = group.Where(s => query.All(word => (s.Text("title") + " " + s.Text("help") + " " + group.Key + " " + string.Join(' ', metadata.SectionFields(s).Select(f => f.Text("label") + " " + f.Text("help")))).Contains(word, StringComparison.OrdinalIgnoreCase))).ToArray();
            if (visible.Length == 0) continue;
            navigation.Children.Add(new TextBlock { Text = group.Key.ToUpperInvariant(), FontSize = 10, LetterSpacing = 1.4, Opacity = .62, FontWeight = FontWeight.SemiBold, Margin = new Thickness(10, 17, 0, 5) });
            foreach (var section in visible)
            {
                var id = section.Text("id");
                var button = new Button { Content = section.Text("title"), Classes = { "nav" } };
                if (id == sectionId) button.Classes.Add("selected");
                button.Click += (_, _) => { SaveScroll(); sectionId = id; RenderNavigation(); RenderPage(); };
                navigation.Children.Add(button);
            }
        }
        if (navigation.Children.Count == 0) navigation.Children.Add(Muted("No matching settings. Try a shorter search."));
    }
    private void RenderPage()
    {
        if (session.Metadata is not { } metadata) return;
        rendering = true;
        var offset = scrollPositions.GetValueOrDefault(ScrollKey, pageScroll.Offset);
        page.Children.Clear();
        var section = metadata.Sections.FirstOrDefault(s => s.Text("id") == sectionId);
        if (section is null) { rendering = false; return; }
        page.Children.Add(new TextBlock { Text = section.Text("title"), Classes = { "heading" } });
        if (section.Text("help") is { Length: > 0 } help) page.Children.Add(Muted(help));
        var inherited = session.Snapshot.Object("inheritance").Bool(sectionId);
        if (section.Text("group") != "" && session.Profile != "Default")
        {
            var inherit = new CheckBox { Content = "Use Default settings for this section", IsChecked = inherited };
            inherit.IsCheckedChanged += async (_, _) => await RunAsync(async () => { await session.MutateAsync("inheritance", new() { ["section"] = sectionId, ["enabled"] = inherit.IsChecked == true }); });
            page.Children.Add(inherit);
        }
        var fields = new StackPanel { Spacing = 16 };
        foreach (var field in metadata.SectionFields(section)) fields.Children.Add(CreateField(field));
        if (fields.Children.Count != 0) page.Children.Add(Card(fields));
        RenderSpecial(section, inherited);
        var commands = metadata.Commands.Where(command => command.Text("section") == sectionId && command.Text("id") != "preview-sound").ToArray();
        if (commands.Length != 0)
        {
            var commandPanel = new WrapPanel();
            foreach (var command in commands)
            {
                var button = ActionButton(command.Text("label"), () => RunCommandAsync(command)); button.Margin = new Thickness(0, 0, 10, 10);
                button.IsEnabled = !command.Bool("runtimeOnly") || session.Role != "offline";
                ToolTip.SetTip(button, !button.IsEnabled ? "Start EVE-X-Preview to use this runtime action." : command.Text("help", command.Text("label")));
                commandPanel.Children.Add(button);
            }
            page.Children.Add(Card(commandPanel));
        }
        var pendingHere = session.Pending.Where(p => p.Key.StartsWith(session.Profile + "|" + sectionId + "|", StringComparison.Ordinal) || p.Key.StartsWith("global|" + sectionId + "|", StringComparison.Ordinal)).ToArray();
        if (pendingHere.Length != 0)
        {
            var feedback = new StackPanel { Spacing = 6 };
            feedback.Children.Add(Muted("Unfinished edits are retained until accepted or explicitly discarded."));
            foreach (var edit in pendingHere.Where(p => p.Error is not null)) feedback.Children.Add(new TextBlock { Text = edit.Label + ": " + edit.Error, Classes = { "error" } });
            page.Children.Add(feedback);
        }
        rendering = false; UpdateStatus();
        Dispatcher.UIThread.Post(() => pageScroll.Offset = offset, DispatcherPriority.Loaded);
    }
    private async Task RunCommandAsync(JsonObject command, JsonObject? arguments = null)
    {
        if (command.Bool("destructive") && !await Dialogs.ConfirmAsync(this, command.Text("label"), "This command runs immediately. Continue?")) return;
        var result = await session.MutateAsync("command", new() { ["name"] = command.Text("id"), ["arguments"] = arguments ?? new JsonObject() });
        await PresentCommandResultAsync(command, result);
    }
    private async Task ApplyAsync()
    {
        SaveScroll();
        if (await session.ApplyAsync())
        {
            notice.Text = "Settings saved. Waiting for the application to restart…";
            for (var attempt = 0; attempt < 16; attempt++)
            {
                await Task.Delay(500);
                try { await ConnectAsync(); return; }
                catch (BridgeException e) when (e.IsDisconnected) { }
            }
            throw new BridgeException("disconnected", "Settings were saved, but the application has not reconnected. Use Reconnect; page and window state are retained.");
        }
        RenderPage();
    }
    private async Task DiscardAsync()
    {
        if (await Dialogs.ConfirmAsync(this, "Discard changes?", "Discard the complete settings draft and all unfinished edits?")) await session.DiscardAsync();
    }
    private async void OnClosing(object? sender, WindowClosingEventArgs args)
    {
        if (closing) return;
        args.Cancel = true;
        if (closePromptOpen) return;
        if (session.HasActiveModal) { error.Text = "Finish or cancel the open dialog before closing settings."; error.IsVisible = true; return; }
        if (busy) { error.Text = "Wait for the current settings operation to finish before closing."; error.IsVisible = true; return; }
        closePromptOpen = true;
        var choice = session.Dirty ? await Dialogs.CloseChoiceAsync(this, session.Pending.Count != 0) : "discard";
        closePromptOpen = false;
        if (choice == "cancel") return;
        if (choice == "apply")
        {
            await RunAsync(ApplyAsync);
            if (session.Dirty || !session.Connected) return;
        }
        try { await session.EndAsync(choice == "discard"); closing = true; Close(); }
        catch (Exception e)
        {
            if (choice == "discard" && !session.Connected) { closing = true; Close(); }
            else ShowError(e);
        }
    }
    private static TextBlock Muted(string text) => new() { Text = text, Classes = { "muted" } };
    private static Border Card(Control child) => new() { Child = child, Classes = { "card" } };
    private string EditKey(string name) => session.Profile + "|" + sectionId + "|" + name;
    internal Task ProtectModalAsync(bool active) => session.SetModalAsync(active);
    private void Stage(string key, string label, string method, JsonObject arguments)
    {
        if (rendering) return;
        session.Stage(key, label, method, arguments);
        if (arguments.Text("id") == "SettingsTheme")
            ThemeController.Apply(arguments.Text("value"), session.Metadata);
        _ = session.ProtectPendingAsync();
    }
}
