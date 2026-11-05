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
    private readonly TextBox search = new() { PlaceholderText = "Search settings", Margin = new Thickness(16, 12), InnerLeftContent = new TextBlock { Text = "⌕", FontSize = 22, Margin = new Thickness(10, 0) } };
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
    private readonly Grid workbench;
    private readonly Border loadingView;
    private readonly bool captionDarkBeforeShow;

    public MainWindow(string installation, SingleInstance instance)
    {
        this.instance = instance;
        session = new EditorSession(new BridgeTransport(installation));
        Width = 1180; Height = 820; MinWidth = 760; MinHeight = 560;
        Title = "EVE-X-Preview · Settings";
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        var root = workbench = new Grid { RowDefinitions = new RowDefinitions("Auto,*,Auto,Auto"),IsVisible=false };
        root.Children.Add(CreateHeader());
        var body = new Grid { ColumnDefinitions = new ColumnDefinitions("242,*") };
        var sidebar = new DockPanel(); DockPanel.SetDock(search, Dock.Top); sidebar.Children.Add(search); sidebar.Children.Add(new ScrollViewer { Content = navigation });
        body.Children.Add(new Border { Child = sidebar, Classes = { "sidebar" }, BorderThickness = new Thickness(0, 1, 1, 1), Opacity = 0.99 });
        pageScroll.Content = page; Grid.SetColumn(pageScroll, 1); body.Children.Add(pageScroll); Grid.SetRow(body, 1); root.Children.Add(body);
        Grid.SetRow(error, 2); root.Children.Add(error);
        var footer = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), Margin = new Thickness(26, 16) };
        var state = new StackPanel { Spacing = 5, VerticalAlignment = VerticalAlignment.Center }; state.Children.Add(draft); state.Children.Add(notice);state.Children.Add(releaseNotice); footer.Children.Add(state);
        releaseNotice.Click+=(_,_)=>{SaveScroll();search.Text="";sectionId="about";RenderNavigation();RenderPage();};
        updatePoll.Tick+=async (_,_)=>await PollUpdatesAsync();
        var actions = new WrapPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center, MaxWidth=Math.Max(400,Width-260) };
        SizeChanged+=(_,_)=>actions.MaxWidth=Math.Max(400,ClientSize.Width-260);
        draft.TextWrapping=Avalonia.Media.TextWrapping.Wrap;
        reconnect = ActionButton("Reconnect", () => ConnectAsync(), false);
        undo = ActionButton("Undo last edit", UndoLastAsync, true);
        accept = ActionButton("Undo this category", UndoCategoryAsync, true);
        discard = ActionButton("Discard all", DiscardAsync, true);
        apply = ActionButton("Apply & restart", ApplyAsync, false); apply.Classes.Add("accent");
        problems.Click+=async (_,_)=>await RunAsync(ShowProblemsAsync,false,false);
        foreach (var button in new[] { reconnect, problems, undo, accept, discard, apply }) { button.Margin = new Thickness(6, 3, 0, 3); actions.Children.Add(button); }
        Grid.SetColumn(actions, 1); footer.Children.Add(actions); Grid.SetRow(footer, 3); root.Children.Add(footer);
        var loading=new StackPanel {Spacing=16,HorizontalAlignment=HorizontalAlignment.Center,VerticalAlignment=VerticalAlignment.Center};
        loading.Children.Add(new TextBlock {Text="EVE-X-Preview",FontSize=24,FontWeight=FontWeight.SemiBold,HorizontalAlignment=HorizontalAlignment.Center});
        loading.Children.Add(new ProgressBar {IsIndeterminate=true,Width=220,Height=4});
        loading.Children.Add(Muted("Loading settings…"));
        loadingView=new Border {Child=loading};Content=new Grid {Children={root,loadingView}};
        search.TextChanged += (_, _) => { RenderNavigation(); RenderPage(); };
        profiles.SelectionChanged += async (_, _) =>
        {
            if (updatingProfiles || profiles.SelectedItem is not string profile || profile == session.Profile) return;
            SaveScroll(); await RunAsync(async () => { await session.MutateAsync("select-profile", new() { ["name"] = profile }); });
        };
        session.Changed += () => Dispatcher.UIThread.Post(UpdateStatus);
        instance.ActivationRequested += ActivateRequested;
        Closing += OnClosing;
        Closed += async (_, _) => {updatePoll.Stop();instance.ActivationRequested -= ActivateRequested; await session.DisposeAsync(); };
        Opened += async (_, _) =>
        {
            if(Program.CaptureDirectory is {} startupDirectory) {
                Directory.CreateDirectory(startupDirectory);
                await Dispatcher.UIThread.InvokeAsync(UpdateLayout,DispatcherPriority.Render);
                using var bitmap=new Avalonia.Media.Imaging.RenderTargetBitmap(new PixelSize((int)Bounds.Width,(int)Bounds.Height));
                bitmap.Render(this);bitmap.Save(Path.Combine(startupDirectory,"startup-loading.png"),Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
            }
            await RunAsync(ConnectAsync);
            workbench.IsVisible=true;loadingView.IsVisible=false;
            if (Program.CaptureDirectory is { } directory && session.Connected)
            {
                try { await CaptureFixturesAsync(directory); }
                catch (Exception error) { System.IO.File.WriteAllText(System.IO.Path.Combine(directory, "capture-error.log"), error.ToString()); closing = true; Close(); }
            }
            else if(Program.CaptureDirectory is {} failedDirectory) {
                File.WriteAllText(Path.Combine(failedDirectory,"startup-error.log"),session.Error??error.Text??"Settings startup failed.");closing=true;Close();
            }
        };
        KeyDown += async (_, args) =>
        {
            if (args.KeyModifiers == KeyModifiers.Control && args.Key == Key.F) { search.Focus(); args.Handled = true; }
            if (args.KeyModifiers == KeyModifiers.Control && args.Key == Key.S && apply.IsEnabled) { args.Handled = true; await RunAsync(ApplyAsync); }
        };
        UpdateStatus();
        captionDarkBeforeShow=NativeWindowFrame.DarkCaption(this);
    }
    private void ActivateRequested() => Dispatcher.UIThread.Post(() => { if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal; Show(); Activate(); });
    private async Task ConnectAsync()
    {
        await session.ConnectAsync();
        await session.ValidateAsync();
        if(!session.Dirty && session.Metadata is {} metadata)ThemeController.RememberTheme(session.Snapshot.Object("values").Object("SettingsTheme").Text("value","system"),metadata);
        if (session.Metadata?.Sections.All(s => s.Text("id") != sectionId) == true) sectionId = session.Metadata.Sections.FirstOrDefault().Text("id");
        RenderNavigation(); RenderPage();
        _=StartBackgroundUpdatesAsync();
    }
    private Button ActionButton(string label, Func<Task> action, bool render = true)
    {
        var button = new Button { Content = label.TrimEnd('.') };
        button.Click += async (_, _) => await RunAsync(action, render, label is not ("Discard all" or "Undo last edit" or "Undo this category" or "Apply & restart"));
        return button;
    }
    private async Task RunAsync(Func<Task> action, bool render = true, bool recordHistory = true)
    {
        if (busy || closing) return;
        busy = true; error.IsVisible = false; UpdateStatus();
        var previousDraft = session.Snapshot.Object("draft").Copy();
        try { await action(); if (recordHistory && previousDraft["config"] is not null && !JsonNode.DeepEquals(previousDraft.Object("config"), session.Snapshot.Object("draft").Object("config"))) editHistory.Push(previousDraft); if (render) { RenderNavigation(); RenderPage(); } }
        catch (Exception e) { ShowError(e); }
        finally { busy = false; UpdateStatus(); }
    }
    private void ShowError(Exception exception)
    {
        error.Text = exception.Message; error.IsVisible = true;
        if (exception is BridgeException { Details: { } details })
        {
            var items = details as JsonArray ?? details.Array("issues");
            foreach (var item in items.Take(8)) error.Text += "\n" + (item is JsonObject ? item.Text("profile") + " / " + item.Text("id") + ": " + item.Text("message") : JsonExtensions.Display(item));
        }
    }
    private void UpdateStatus()
    {
        var pending = session.Pending.Count;
        draft.Text = pending != 0 ? $"{pending} invalid or validating field{(pending == 1 ? "" : "s")}" : session.Dirty ? "Unapplied changes" : "All changes applied";
        notice.Text = busy ? "Working" : pending != 0 ? "Correct the highlighted fields before applying." : activity;
        if (!busy && session.Pending.FirstOrDefault(p => p.Error is not null) is { } invalid)
            notice.Text = "Invalid: " + invalid.Profile + " / " + invalid.Label;
        else if (!busy && session.Validation.Array("issues").FirstOrDefault(i => i.Text("severity") == "error") is { } issue)
            notice.Text = "Invalid: " + issue.Text("profile") + " / " + (session.Metadata?.Fields.FirstOrDefault(f => f.Text("id") == issue.Text("id")).Text("label", issue.Text("id")) ?? issue.Text("id"));
        var issues=CurrentProblems();var blocking=issues.Any(p=>p.Severity=="error");
        draft.Foreground=blocking?Brushes.OrangeRed:null;notice.Foreground=blocking?Brushes.OrangeRed:null;
        notice.MaxHeight=48;ToolTip.SetTip(notice,notice.Text);
        problems.Content=issues.Count==0?"Problems":"Problems ("+issues.Count+")";
        problems.IsEnabled=!busy&&session.Connected&&issues.Count>0;
        apply.IsEnabled=!busy&&session.Connected&&session.Dirty&&pending==0&&!blocking;
        accept.IsEnabled = !busy && session.Connected && session.Dirty;
        discard.IsEnabled = !busy && session.Connected && session.Dirty;
        undo.IsEnabled = !busy && session.Connected && (editHistory.Count != 0 || pending != 0);
        reconnect.IsVisible = !session.Connected; reconnect.IsEnabled = !busy;
        profiles.IsEnabled = !busy && session.Connected;
        page.IsEnabled = !busy && session.Connected;
        RefreshAboutControls();
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
        bool Matches(JsonObject section) => query.All(word => (section.Text("title") + " " + section.Text("help") + " " + section.Text("page") + " " + string.Join(' ', metadata.SectionFields(section).Select(f => f.Text("label") + " " + f.Text("help") + " " + string.Join(' ',f.Array("choices").Select(c=>c.Text("label")))))).Contains(word,StringComparison.OrdinalIgnoreCase));
        var matches = metadata.Sections.Where(Matches).ToArray();
        if(matches.Length != 0 && !matches.Any(s=>s.Text("id")==sectionId)) { SaveScroll(); sectionId=matches[0].Text("id"); }
        foreach (var group in metadata.Sections.GroupBy(s => s.Text("page")))
        {
            var visible = group.Where(Matches).ToArray();
            if (visible.Length == 0) continue;
            navigation.Children.Add(new TextBlock { Text = group.Key.ToUpperInvariant(), FontSize = 10, LetterSpacing = 1.4, Opacity = .62, FontWeight = FontWeight.SemiBold, Margin = new Thickness(10, 17, 0, 5) });
            foreach (var section in visible)
            {
                var id = section.Text("id");
                var button = new Button { Content = SearchText(section.Text("title")), Classes = { "nav" } };
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
        profileBaselines.TryAdd(session.Profile, session.Snapshot.Object("draft").Object("config").Copy());
        var offset = scrollPositions.GetValueOrDefault(ScrollKey, pageScroll.Offset);
        page.Children.Clear(); fieldFeedback.Clear();
        var section = metadata.Sections.FirstOrDefault(s => s.Text("id") == sectionId);
        if (section is null) { rendering = false; return; }
        var titleBlock = SearchText(section.Text("title")); titleBlock.Classes.Add("heading"); page.Children.Add(titleBlock);
        if (section.Text("help") is { Length: > 0 } help) page.Children.Add(SearchText(help));
        var inherited = session.Snapshot.Object("inheritance").Bool(sectionId);
        if (section.Text("group") != "" && session.Profile != "Default")
        {
            var inherit = new CheckBox { Content = "Use Default settings for this section", IsChecked = inherited };
            inherit.IsCheckedChanged += async (_, _) => await RunAsync(async () => { await session.MutateAsync("inheritance", new() { ["section"] = sectionId, ["enabled"] = inherit.IsChecked == true }); });
            page.Children.Add(inherit);
        }
        var fields = new StackPanel { Spacing = 16 };
        foreach (var field in metadata.SectionFields(section)) {
            if(sectionId=="events" && field.Text("id")=="shootingInterval")continue;
            var editor=CreateField(field,sectionId=="sounds");
            fields.Children.Add(editor);
        }
        if (fields.Children.Count != 0) page.Children.Add(Card(fields));
        RenderSpecial(section, inherited);
        var commands = metadata.Commands.Where(command => command.Text("section") == sectionId && command.Text("id") != "preview-sound" && sectionId!="about").ToArray();
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

            foreach (var edit in pendingHere.Where(p => p.Error is not null)) feedback.Children.Add(new TextBlock { Text = edit.Label + ": " + edit.Error, Classes = { "error" } });
            page.Children.Add(feedback);
        }
        rendering = false; UpdateStatus();
        Dispatcher.UIThread.Post(() => pageScroll.Offset = offset, DispatcherPriority.Loaded);
    }
    private async Task RunCommandAsync(JsonObject command, JsonObject? arguments = null)
    {
        if (command.Bool("destructive") && !await Dialogs.ConfirmAsync(this, command.Text("label"), "This command runs immediately. Continue?")) return;
        var result = await session.ReadAsync("command", new() { ["name"] = command.Text("id"), ["arguments"] = arguments ?? new JsonObject() });
        await PresentCommandResultAsync(command, result);
    }
    private async Task ApplyAsync()
    {
        SaveScroll();
        var restart = await session.ApplyAsync();
        if(session.Metadata is {} metadata)ThemeController.RememberTheme(session.Snapshot.Object("values").Object("SettingsTheme").Text("value","system"),metadata);
        editHistory.Clear(); profileBaselines.Clear();
        if (restart)
        {
            notice.Text = "Settings saved. Waiting for the application to restart";
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
        if (!await Dialogs.ConfirmAsync(this, "Discard changes?", "Discard all unapplied changes?")) return;
        foreach (var timer in validationTimers.Values) timer.Cancel();
        await liveEdits.WaitAsync();
        try { await session.DiscardAsync(); editHistory.Clear(); profileBaselines.Clear(); }
        finally { liveEdits.Release(); }
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
        foreach (var timer in validationTimers.Values) timer.Cancel();
        await liveEdits.WaitAsync();
        try { await session.EndAsync(choice == "discard"); closing = true; Close(); }
        catch (Exception e)
        {
            if (choice == "discard" && !session.Connected) { closing = true; Close(); }
            else ShowError(e);
        }
        finally { liveEdits.Release(); }
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
        QueueValidation(key);
    }
}
