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
    private readonly Button clearSearch = new() { Content = new PathIcon { Data = Geometry.Parse("M0,1 L1,0 L4,3 L7,0 L8,1 L5,4 L8,7 L7,8 L4,5 L1,8 L0,7 L3,4 Z"), Width = 9, Height = 9 }, HorizontalContentAlignment = HorizontalAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center, Width = 26, Height = 26, MinHeight = 0, Padding = new Thickness(0), Margin = new Thickness(2, 0, 5, 0), Background = Brushes.Transparent, BorderThickness = new Thickness(0), IsVisible = false };
    private readonly Border connectionBanner = new() { IsVisible = false, Padding = new Thickness(26, 12), Background = new SolidColorBrush(Color.Parse("#553717")) };
    private readonly TextBlock connectionMessage = new() { TextWrapping = TextWrapping.Wrap, Foreground = Brushes.White };
    private readonly DispatcherTimer connectionPoll = new() { Interval = TimeSpan.FromSeconds(2) };
    private string? defaultEditProfile, defaultEditSection, defaultEditGroup;
    private string SelectedProfile => defaultEditProfile ?? session.Profile;
    private readonly TextBlock draft = new() { FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock error = new() { Classes = { "error" }, IsVisible = false, Margin = new Thickness(28, 8) };
    private readonly TextBlock notice = Muted("");
    private readonly Button apply, discard, accept, undo, reconnect;
    private readonly Dictionary<string, Vector> scrollPositions = new();
    private readonly Dictionary<string, string> selections = new();
    private string sectionId = "appearance";
    private bool busy, controlsLocked, closing, closePromptOpen, updatingProfiles, rendering, suppressSearchRefresh;
    private int navigationRequest;
    private bool reconnecting;
    private int automaticReconnectFailures;
    private string? renderedScrollKey;
    private Vector requestedScroll;
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
        var root = workbench = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,*,Auto,Auto"),IsVisible=false };
        root.Children.Add(new Border { Child = CreateHeader(), Classes = { "chrome" }, BorderThickness = new Thickness(0,0,0,1) });
        connectionBanner.Child = connectionMessage; Grid.SetRow(connectionBanner, 1); root.Children.Add(connectionBanner);
        search.InnerRightContent = clearSearch;
        Avalonia.Automation.AutomationProperties.SetName(clearSearch, "Clear search");
        ToolTip.SetTip(clearSearch, "Clear search");
        clearSearch.Click += (_, _) => { search.Text = ""; search.Focus(); };
        var body = new Grid { ColumnDefinitions = new ColumnDefinitions("242,*") };
        var sidebar = new DockPanel(); DockPanel.SetDock(search, Dock.Top); sidebar.Children.Add(search); sidebar.Children.Add(new ScrollViewer { Content = navigation });
        body.Children.Add(new Border { Child = sidebar, Classes = { "sidebar" }, BorderThickness = new Thickness(0, 0, 1, 0) });
        pageScroll.Content = page; Grid.SetColumn(pageScroll, 1); body.Children.Add(pageScroll); Grid.SetRow(body, 2); root.Children.Add(body);
        Grid.SetRow(error, 3); root.Children.Add(error);
        var footer = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), Margin = new Thickness(26, 16) };
        var state = new StackPanel { Spacing = 5, VerticalAlignment = VerticalAlignment.Center }; state.Children.Add(draft); state.Children.Add(notice);state.Children.Add(releaseNotice); footer.Children.Add(state);
        releaseNotice.Click+=async (_,_)=> {try {await NavigateToAsync("about",true);}catch(Exception e) {ShowError(e);}};
        updatePoll.Tick+=async (_,_)=>await PollUpdatesAsync();
        var actions = new WrapPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center, MaxWidth=Math.Max(400,Width-260) };
        SizeChanged+=(_,_)=>actions.MaxWidth=Math.Max(400,ClientSize.Width-260);
        draft.TextWrapping=Avalonia.Media.TextWrapping.Wrap;
        reconnect = ActionButton("Reconnect", () => AutoReconnectAsync(), false);
        undo = ActionButton("Undo last edit", UndoLastAsync, true);
        accept = ActionButton("Undo this category", UndoCategoryAsync, true);
        discard = ActionButton("Discard all", DiscardAsync, true);
        apply = ActionButton("Apply & restart", ApplyAsync, false); apply.Classes.Add("accent");
        problems.Click+=async (_,_)=>await RunAsync(ShowProblemsAsync,false,false);
        foreach (var button in new[] { reconnect, problems, undo, accept, discard, apply }) { button.Margin = new Thickness(6, 3, 0, 3); actions.Children.Add(button); }
        Grid.SetColumn(actions, 1); footer.Children.Add(actions); var footerBand = new Border { Child = footer, Classes = { "chrome" }, BorderThickness = new Thickness(0,1,0,0) }; Grid.SetRow(footerBand, 4); root.Children.Add(footerBand);
        var loading=new StackPanel {Spacing=16,HorizontalAlignment=HorizontalAlignment.Center,VerticalAlignment=VerticalAlignment.Center};
        loading.Children.Add(new TextBlock {Text="EVE-X-Preview",FontSize=24,FontWeight=FontWeight.SemiBold,HorizontalAlignment=HorizontalAlignment.Center});
        loading.Children.Add(new ProgressBar {IsIndeterminate=true,Width=220,Height=4});
        loading.Children.Add(Muted("Loading settings…"));
        loadingView=new Border {Child=loading};Content=new Grid {Children={root,loadingView}};
        search.TextChanged += async (_, _) => { clearSearch.IsVisible = !string.IsNullOrEmpty(search.Text); if(suppressSearchRefresh)return; SaveScroll(); try {if(defaultEditProfile is not null)await EndDefaultEditingAsync(); RenderNavigation(); RenderPage();}catch(Exception e) {ShowError(e);} };
        profiles.SelectionChanged += async (_, _) =>
        {
            if (updatingProfiles || profiles.SelectedItem is not string profile || profile == SelectedProfile) return;
            SaveScroll(); await RunAsync(async () => { defaultEditProfile = defaultEditSection = defaultEditGroup = null; await session.MutateAsync("select-profile", new() { ["name"] = profile }); },exclusive:true);
        };
        session.Changed += () => Dispatcher.UIThread.Post(UpdateStatus);
        instance.ActivationRequested += ActivateRequested;
        Closing += OnClosing;
        connectionPoll.Tick += async (_, _) => {
            if(closing || busy || reconnecting || session.HasActiveModal || liveEdits.CurrentCount == 0)return;
            if(!session.Connected) {if(automaticReconnectFailures==0)await AutoReconnectAsync();return;}
            try { await session.ReadAsync("snapshot");if(sectionId=="diagnostics")await RefreshPreviewStatusAsync(); } catch(Exception) { UpdateStatus();await AutoReconnectAsync(); }
        };
        Closed += async (_, _) => {connectionPoll.Stop();updatePoll.Stop();instance.ActivationRequested -= ActivateRequested; await session.DisposeAsync(); };
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
        automaticReconnectFailures=0;
        connectionPoll.Start();
        if(!session.Dirty && session.Metadata is {} metadata)ThemeController.RememberTheme(session.Snapshot.Object("values").Object("SettingsTheme").Text("value","system"),metadata);
        if (session.Metadata?.Sections.All(s => s.Text("id") != sectionId) == true) sectionId = session.Metadata.Sections.FirstOrDefault().Text("id");
        RenderNavigation(); RenderPage();
        foreach(var edit in session.Pending.ToArray())QueueValidation(edit.Key);
        _=StartBackgroundUpdatesAsync();
    }
    private Button ActionButton(string label, Func<Task> action, bool render = true)
    {
        var button = new Button { Content = label.TrimEnd('.') };
        button.Click += async (_, _) => await RunAsync(action, render, label is not ("Discard all" or "Undo last edit" or "Undo this category" or "Apply & restart"),label is "Apply & restart" or "Reconnect" or "Discard all");
        return button;
    }
    private async Task RunAsync(Func<Task> action, bool render = true, bool recordHistory = true, bool exclusive = false)
    {
        if (closing) return;
        if(render)SaveScroll();
        refreshAfterActions |= render;
        busy = true; controlsLocked |= exclusive; error.IsVisible = false; UpdateStatus();
        try {
            await actionQueue.RunAsync(async () => {
                var previousDraft = session.Snapshot.Object("draft").Copy();
                try {
                    await action();
                    if(session.Connected && render)await session.ValidateAsync();
                    if (recordHistory && previousDraft["config"] is not null && !JsonNode.DeepEquals(previousDraft.Object("config"), session.Snapshot.Object("draft").Object("config"))) editHistory.Push(previousDraft);
                }
                catch (Exception e) { ShowError(e); }
            });
        }
        finally {
            busy = actionQueue.PendingCount != 0;
            if (!busy) {
                controlsLocked = false;
                if (refreshAfterActions) { refreshAfterActions = false; RenderNavigation(); RenderPage(); }
            }
            UpdateStatus();
        }
    }
    private readonly EditorActionQueue actionQueue = new();
    private bool refreshAfterActions;
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
        var changed = session.ChangeCount;
        draft.Text = changed != 0 ? $"{changed} setting{(changed == 1 ? "" : "s")} changed, apply to save" : "All changes applied";
        notice.Text = busy ? "Working" : pending != 0 ? "Correct the highlighted fields before applying." : activity;
        if (!busy && session.Pending.FirstOrDefault(p => p.Error is not null) is { } invalid)
            notice.Text = "Invalid: " + invalid.Profile + " / " + invalid.Label;
        else if (!busy && session.Validation.Array("issues").FirstOrDefault(i => i.Text("severity") == "error") is { } issue)
            notice.Text = "Invalid: " + issue.Text("profile") + " / " + (session.Metadata?.Fields.FirstOrDefault(f => f.Text("id") == issue.Text("id")).Text("label", issue.Text("id")) ?? issue.Text("id"));
        var issues=CurrentProblems();var blocking=issues.Any(p=>p.Severity=="error");
        if(blocking) {draft.Foreground=Brushes.OrangeRed;notice.Foreground=Brushes.OrangeRed;}
        else {draft.ClearValue(TextBlock.ForegroundProperty);notice.ClearValue(TextBlock.ForegroundProperty);}
        notice.MaxHeight=48;ToolTip.SetTip(notice,notice.Text);
        problems.Content=issues.Count==0?"Problems":"Problems ("+issues.Count+")";
        problems.IsEnabled=!busy&&session.Connected&&issues.Count>0;
        apply.IsEnabled=!busy&&session.Connected&&session.Dirty&&pending==0&&!blocking;
        accept.IsEnabled = !busy && session.Connected && session.Dirty;
        discard.IsEnabled = !busy && session.Connected && session.Dirty;
        undo.IsEnabled = !busy && session.Connected && (editHistory.Count != 0 || pending != 0);
        reconnect.IsVisible = !session.Connected; reconnect.IsEnabled = !busy;
        connectionBanner.IsVisible = !session.Connected;
        connectionMessage.Text = (session.AwaitingRestart && automaticReconnectFailures==0 ? "Settings saved. Waiting for the application to restart and reconnecting automatically. Your selected profile is retained." : reconnecting ? "Reconnecting automatically. Your draft and unfinished entries are retained." : "Settings disconnected. Editing is paused; your draft and unfinished entries are retained. Click Reconnect to continue.") + (session.Error is { Length: > 0 } reason ? "\n" + reason : "");
        profiles.IsEnabled = !busy && !reconnecting && session.Connected;
        search.IsEnabled = !busy && !reconnecting && session.Connected;
        navigation.IsEnabled = !busy;
        page.IsEnabled = !controlsLocked && !reconnecting && session.Connected;
        RefreshAboutControls();
        if (!rendering)
        {
            updatingProfiles = true;
            var names = session.Snapshot.Array("profiles").Select(item => item is JsonObject ? item.Text("name") : JsonExtensions.Display(item)).ToArray();
            if (!names.SequenceEqual((profiles.ItemsSource as string[]) ?? [])) profiles.ItemsSource = names;
            profiles.SelectedItem = SelectedProfile; updatingProfiles = false;
        }
        if (session.Snapshot.Object("values")["SettingsTheme"] is { } theme)
        {
            var requested = session.Pending.FirstOrDefault(edit => edit.Arguments.Text("id") == "SettingsTheme")?.Arguments.Text("value") ?? theme.Text("value");
            ThemeController.Apply(captureTheme ?? requested, session.Metadata);
        }
    }
    private string ScrollKey => session.Profile + "|" + sectionId;
    private void SaveScroll() { if(renderedScrollKey is {} key)scrollPositions[key] = pageScroll.Offset; }
    private void RestoreLoadedScroll(int generation) => Dispatcher.UIThread.Post(()=> { if(generation==pageGeneration)pageScroll.Offset=requestedScroll; },DispatcherPriority.Loaded);
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
                button.Click += async (_, _) => { try { await NavigateToAsync(id); } catch(Exception e) { ShowError(e); } };
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
        var offset = renderedScrollKey==ScrollKey ? pageScroll.Offset : scrollPositions.GetValueOrDefault(ScrollKey, default);
        requestedScroll=offset;renderedScrollKey=ScrollKey;
        page.Children.Clear(); fieldFeedback.Clear();
        var section = metadata.Sections.FirstOrDefault(s => s.Text("id") == sectionId);
        if (section is null) { rendering = false; return; }
        var titleBlock = SearchText(section.Text("title")); titleBlock.Classes.Add("heading"); page.Children.Add(titleBlock);
        if (section.Text("help") is { Length: > 0 } help) page.Children.Add(SearchText(help));
        var inherited = session.Snapshot.Object("inheritance").Bool(sectionId);
        RenderSources(section);
        RenderDependencySummary(section);
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
                var button = ActionButton(command.Text("label"), () => RunCommandAsync(command),false);
                if(command.Text("id")=="toggle-live-previews") { previewPauseButton=button;button.Content="Pause live previews";_ = RefreshPreviewStatusAsync(); } button.Margin = new Thickness(0, 0, 10, 10);
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
        RestoreLoadedScroll(pageGeneration);
    }
    private async Task NavigateToAsync(string id, bool clearQuery = false)
    {
        var request=++navigationRequest;
        SaveScroll();
        if(clearQuery) {suppressSearchRefresh=true;search.Text="";suppressSearchRefresh=false;}
        if(defaultEditProfile is not null && id != defaultEditSection)await EndDefaultEditingAsync();
        if(request!=navigationRequest)return;
        sectionId = id; RenderNavigation(); RenderPage();
    }
    private async Task EndDefaultEditingAsync()
    {
        if(defaultEditProfile is not {} selected)return;
        await session.MutateAsync("select-profile",new(){["name"]=selected});
        defaultEditProfile = defaultEditSection = defaultEditGroup = null;
    }
    private async Task SetInheritanceAsync(string group, bool enabled)
    {
        if(enabled) foreach(var edit in session.Pending.Where(p=>p.Profile==session.Profile &&
            (p.Key.Split('|').ElementAtOrDefault(1)==group || session.Metadata?.Field(p.Arguments.Text("id"))?.Text("group")==group)).ToArray()) {
            if(validationTimers.TryGetValue(edit.Key,out var timer))timer.Cancel();
            await session.DiscardPendingAsync(edit.Key);
        }
        await session.MutateAsync("inheritance",new(){["section"]=group,["enabled"]=enabled});
    }
    private void RenderSources(JsonObject section)
    {
        if(defaultEditProfile is not null) {
            var context = new StackPanel { Spacing = 8 };
            context.Children.Add(Muted("Editing Default · " + session.Metadata!.Sections.FirstOrDefault(s=>s.Text("id")==defaultEditGroup).Text("title","this category") + ". Changes affect profiles inheriting it. Selected profile: " + defaultEditProfile));
            context.Children.Add(ActionButton("Done editing Default",EndDefaultEditingAsync));
            page.Children.Add(Card(context));return;
        }
        var groups = new[] { section.Text("group") }.Concat(session.Metadata!.SectionFields(section).Where(f=>!f.Bool("global")).Select(f=>f.Text("group")));
        if(sectionId == "quick-group")groups = groups.Append("mouse-actions");
        if(sectionId == "sounds")groups = groups.Append("events");
        foreach(var group in groups.Where(g=>g!="").Distinct()) {
            var label = session.Metadata.Sections.FirstOrDefault(s=>s.Text("id")==group).Text("title",group);
            if(session.Profile == "Default") { page.Children.Add(Muted("Source: Default profile · " + label));continue; }
            var source = new StackPanel { Spacing = 8 };
            source.Children.Add(Muted("Settings source: " + label + (sectionId=="sounds" && group=="events"?" · event sound assignments":"")));
            var inherited = session.Snapshot.Object("inheritance").Bool(group);
            if(inherited)source.Children.Add(Muted(sectionId=="sounds" && group=="events" ? "Event sound assignments are inherited from Default and cannot be changed here. Edit Default here, or customize Event alerts for this profile." : "This category is inherited from Default. Edit Default here, or customize it to enable editing for this profile."));
            var choice = new ComboBox { ItemsSource = new[] { "Use Default profile settings", "Customize for this profile" }, SelectedIndex = inherited ? 0 : 1, MinWidth = 260, HorizontalAlignment = HorizontalAlignment.Left };
            choice.SelectionChanged += async (_,_) => {
                var enabled = choice.SelectedIndex == 0;
                if(enabled == inherited)return;
                inherited = enabled;
                await RunAsync(()=>SetInheritanceAsync(group,enabled));
            };
            source.Children.Add(choice);
            if(inherited)source.Children.Add(ActionButton(group==sectionId?"Edit Default here":sectionId=="sounds" && group=="events"?"Edit Default sound assignments here":"Edit Default " + label,async()=> {
                SaveScroll();defaultEditProfile = session.Profile;defaultEditGroup = group;defaultEditSection = sectionId=="sounds" && group=="events" ? "sounds" : group;sectionId = defaultEditSection;
                try {await session.MutateAsync("select-profile",new(){["name"]="Default"});}
                catch {defaultEditProfile = defaultEditSection = defaultEditGroup = null;throw;}
            }));
            page.Children.Add(source);
        }
    }
    private async Task RunCommandAsync(JsonObject command, JsonObject? arguments = null)
    {
        if (command.Bool("destructive") && !await Dialogs.ConfirmAsync(this, command.Text("label"), "This command runs immediately. Continue?")) return;
        var result = await session.MutateAsync("command", new() { ["name"] = command.Text("id"), ["arguments"] = arguments ?? new JsonObject() });
        await PresentCommandResultAsync(command, result);
        if(command.Text("id")=="toggle-live-previews")SetPreviewStatus(result);
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
            automaticReconnectFailures=0;
            await Task.Delay(750);
            await AutoReconnectAsync();
            return;
        }
        RenderPage();
    }

    private async Task AutoReconnectAsync()
    {
        if(reconnecting || closing)return;
        reconnecting=true;UpdateStatus();
        try {
            await session.ReconnectAsync(attempt=> {notice.Text=$"Reconnecting ({attempt}/5)…";return Task.CompletedTask;});
            automaticReconnectFailures=0;
            await ConnectPresentationAsync();
        } catch(Exception e) {automaticReconnectFailures++;ShowError(e);}
        finally {reconnecting=false;UpdateStatus();}
    }
    private Task ConnectPresentationAsync()
    {
        error.IsVisible=false;
        connectionPoll.Start();
        RenderNavigation();RenderPage();
        foreach(var edit in session.Pending.ToArray())QueueValidation(edit.Key);
        _=StartBackgroundUpdatesAsync();
        return Task.CompletedTask;
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
