using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using EveX.Settings.Bridge;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    private int pageGeneration;
    private void RenderSpecial(JsonObject section, bool inherited)
    {
        var generation = ++pageGeneration;
        var id = section.Text("id");
        if (id == "profiles") RenderProfiles();
        if (id == "inheritance") RenderInheritance();
        if (id == "events") RenderEvents(generation);
        if (id == "sounds") RenderSounds(generation);
        if (id == "about") RenderAbout(generation);
        if (id is "character-groups" or "app-groups")
        {
            var host = new StackPanel { Spacing = 12 }; page.Children.Add(Card(host));
            LoadGroupEditor(host, id == "character-groups" ? "characters" : "apps", inherited, generation);
        }
        else
        {
            foreach (var table in session.Metadata!.Tables.Where(table => table.Text("section") == id))
            {
                var host = new StackPanel { Spacing = 12 }; page.Children.Add(Card(host));
                LoadRowsEditor(host, table.Text("kind"), "", inherited, generation);
            }
        }

    }
    private void RenderProfiles()
    {
        var panel = new StackPanel { Spacing = 14 };
        var active = session.Snapshot.Array("profiles").FirstOrDefault(profile => profile.Bool("active"));
        panel.Children.Add(Muted("Active profile: " + active.Text("name", session.Profile) + "\nSelecting a profile changes the draft you are editing. Activation takes effect on Apply."));
        var buttons = new WrapPanel();
        buttons.Children.Add(ActionButton("Create profile", CreateProfileAsync));
        var rename = ActionButton("Rename", async () =>
        {
            await Dialogs.NameActionAsync(this, "Rename profile", session.Profile, async name => { await session.MutateAsync("profile", new() { ["action"] = "rename", ["name"] = name }); });
        }); rename.IsEnabled = session.Profile != "Default"; rename.Margin = new Thickness(8, 0, 0, 0); buttons.Children.Add(rename);
        var remove = ActionButton("Remove profile", async () =>
        {
            if (await Dialogs.ConfirmAsync(this, "Remove profile?", "Remove “" + session.Profile + "” from the draft? Undo restores it.")) await session.MutateAsync("profile", new() { ["action"] = "delete" });
        }); remove.IsEnabled = session.Profile != "Default"; remove.Margin = new Thickness(8, 0, 0, 0); buttons.Children.Add(remove);
        var activate = ActionButton("Activate on Apply", async () => { await session.MutateAsync("profile", new() { ["action"] = "activate" }); }); activate.Margin = new Thickness(8, 0, 0, 0); buttons.Children.Add(activate);
        panel.Children.Add(buttons);
        if (session.Profile == "Default") panel.Children.Add(Muted("Default is built in. Its settings are the source for inherited sections and it cannot be renamed or removed."));
        page.Children.Add(Card(panel));
    }
    private async Task CreateProfileAsync()
    {
        var content = new StackPanel { Spacing = 10 };
        var name = new TextBox { PlaceholderText = "Profile name" };
        var copy = new CheckBox { Content = "Copy current profile settings", IsChecked = true };
        content.Children.Add(name); content.Children.Add(copy);
        content.Children.Add(Muted("Choose sections that should inherit from Default:"));
        var flags = session.Snapshot.Object("creationInheritance").Copy();
        var (choices, checks) = InheritanceChoices(flags);
        var presets = new WrapPanel();
        foreach(var (label, mode) in new[] { ("Select all", "all"), ("Select none", "none"), ("Use current profile inheritance", "current") }) {
            var button = new Button { Content = label, Margin = new Thickness(0,0,8,8) };
            button.Click += (_,_) => {
                foreach(var (id, check) in checks) {
                    var enabled = mode == "all" || mode == "current" && session.Snapshot.Object("inheritance").Bool(id);
                    flags[id] = enabled;check.IsChecked = enabled;
                }
            };
            presets.Children.Add(button);
        }
        content.Children.Add(presets);
        await Dialogs.ChoiceAsync(this, "Create profile", choices, [("Create", "create"), ("Cancel", "cancel")], async choice =>
        {
            if (choice == "create") await session.MutateAsync("profile", new() { ["action"] = "add", ["name"] = name.Text, ["copyCurrent"] = copy.IsChecked == true, ["inheritance"] = flags.DeepClone() });
        }, fixedContent: content);
    }
    private (StackPanel Panel, Dictionary<string,CheckBox> Checks) InheritanceChoices(JsonObject flags, Func<string,bool,Task>? save = null)
    {
        var panel = new StackPanel { Spacing = 8 };
        var checks = new Dictionary<string,CheckBox>();
        foreach(var category in session.Metadata!.Sections.Where(s=>s.Text("group")!="").GroupBy(s=>s.Text("page"))) {
            panel.Children.Add(new Border { Height = 1, Background = ThemeController.Line, Margin = new Thickness(0,12,0,4) });
            panel.Children.Add(new TextBlock { Text = category.Key, FontWeight = FontWeight.SemiBold });
            foreach(var section in category) {
                var id = section.Text("id");
                var check = new CheckBox { Content = section.Text("title"), IsChecked = flags.Bool(id), Tag = id };
                Avalonia.Automation.AutomationProperties.SetName(check,"Inherit " + section.Text("title"));
                check.IsCheckedChanged += async (_,_) => {
                    var enabled = check.IsChecked == true;
                    if(flags.Bool(id)==enabled)return;
                    flags[id] = enabled;
                    if(save is not null)await RunAsync(()=>save(id,enabled));
                };
                checks[id] = check;panel.Children.Add(check);
            }
        }
        return (panel,checks);
    }
    private void RenderInheritance()
    {
        var template = session.Profile == "Default";
        var panel = new StackPanel { Spacing = 10 };
        panel.Children.Add(Muted(template ? "Checked categories are inherited from Default by new profiles. Existing profiles keep their current sources." : "Checked categories inherit from Default. Unchecked categories use this profile's stored values."));
        var sections = session.Metadata!.Sections.Where(s => s.Text("group") != "").ToArray();
        var flags = session.Snapshot.Object(template ? "creationInheritance" : "inheritance").Copy();
        var presets = new WrapPanel();
        foreach (var (label, mode) in new[] { ("Inherit all", "all"), ("Customize all", "none"), ("Use new-profile template", "template") }) {
            if(template && mode == "template")continue;
            var button = ActionButton(label, async () => {
                var values = new JsonObject();
                foreach(var section in sections) values[section.Text("id")] = mode == "all" || mode == "template" && session.Snapshot.Object("creationInheritance").Bool(section.Text("id"));
                if(template) await session.MutateAsync("creation-inheritance", new() { ["flags"] = values });
                else foreach(var section in sections) await SetInheritanceAsync(section.Text("id"), values.Bool(section.Text("id")));
            });
            button.Margin = new Thickness(0,0,8,8);presets.Children.Add(button);
        }
        panel.Children.Add(presets);
        var (choices, _) = InheritanceChoices(flags,async (id,enabled) => {
            if(template)await session.MutateAsync("creation-inheritance",new(){["flags"]=flags.DeepClone()});
            else await SetInheritanceAsync(id,enabled);
        });
        panel.Children.Add(choices);
        page.Children.Add(Card(panel));
    }
    private async void LoadGroupEditor(StackPanel host, string kind, bool inherited, int generation)
    {
        try
        {
            host.Children.Add(Muted("Loading groups"));
            var tableKind = kind == "characters" ? "group-members" : "app-group-members";
            var data = await session.ReadAsync("rows", new() { ["kind"] = tableKind });
            if (generation != pageGeneration) return;
            host.Children.Clear();
            var names = data.Array("groups").Select(item => item is JsonObject ? item.Text("name") : JsonExtensions.Display(item)).ToArray();
            var selectionKey = session.Profile + "|group|" + kind;
            var selected = selections.GetValueOrDefault(selectionKey, names.FirstOrDefault() ?? "");
            if (!names.Contains(selected)) selected = names.FirstOrDefault() ?? "";
            selections[selectionKey] = selected;
            var selector = new ComboBox { ItemsSource = names, SelectedItem = selected, HorizontalAlignment = HorizontalAlignment.Stretch };
            selector.SelectionChanged += (_, _) => { if (selector.SelectedItem is string name) { selections[selectionKey] = name; RenderPage(); } }; host.Children.Add(selector);
            var buttons = new WrapPanel();
            buttons.Children.Add(ActionButton("New group", async () => { await Dialogs.NameActionAsync(this, "Create group", "", async name => { await session.MutateAsync("group", new() { ["kind"] = kind, ["action"] = "add", ["name"] = name }); selections[selectionKey] = name; }); }));
            if (selected != "")
            {
                var rename = ActionButton("Rename", async () => { await Dialogs.NameActionAsync(this, "Rename group", selected, async name => { await session.MutateAsync("group", new() { ["kind"] = kind, ["action"] = "rename", ["name"] = selected, ["newName"] = name }); selections[selectionKey] = name; }); }); rename.Margin = new Thickness(8, 0, 0, 0); buttons.Children.Add(rename);
                var remove = ActionButton("Remove group", async () => { if (await Dialogs.ConfirmAsync(this, "Remove group?", "Remove “" + selected + "” from the draft? Undo restores it and its members.")) await session.MutateAsync("group", new() { ["kind"] = kind, ["action"] = "delete", ["name"] = selected }); }); remove.Margin = new Thickness(8, 0, 0, 0); buttons.Children.Add(remove);
            }
            buttons.IsEnabled = !inherited; host.Children.Add(buttons);
            if (selected == "") { host.Children.Add(Muted("Create a group to add members and switching shortcuts.")); return; }
            var selectedData = await session.ReadAsync("rows", new() { ["kind"] = tableKind, ["group"] = selected });
            if (generation != pageGeneration) return;
            foreach (var shortcut in selectedData.Array("groupShortcuts").OfType<JsonObject>())
            {
                var key = shortcut.Text("key"); var editKey = EditKey("group|" + kind + "|" + selected + "|" + key);
                host.Children.Add(new TextBlock { Text = shortcut.Text("label"), FontWeight = FontWeight.SemiBold });
                var value = session.GetPending(editKey)?.Arguments.Text("value") ?? shortcut.Text("value");
                var control = new HotkeyEditor(value, text => Stage(editKey, shortcut.Text("label"), "group", new() { ["kind"] = kind, ["action"] = "shortcut", ["name"] = selected, ["key"] = key, ["value"] = text })); control.IsEnabled = !inherited; host.Children.Add(control);
            }
            var rowHost = new StackPanel { Spacing = 12 }; host.Children.Add(rowHost); BuildRowsEditor(rowHost, tableKind, selected, selectedData, inherited);
        }
        catch (Exception e) { if (generation == pageGeneration) { host.Children.Clear(); host.Children.Add(new TextBlock { Text = e.Message, Classes = { "error" } }); } }
    }
    private void RenderEvents(int generation) {
        var table=new StackPanel {Spacing=2};page.Children.Add(Card(table));LoadEventOverview(table,generation);
    }
    private void RenderSounds(int generation) {
        var host=new StackPanel {Spacing=10};page.Children.Add(host);LoadSoundOverview(host,generation);
    }
    private async Task AssignSoundAsync(string path, JsonArray assigned)
    {
        var content = new StackPanel { Spacing = 8 }; content.Children.Add(Muted("Checked events use this sound. Unchecking only removes this sound from that event; other assignments remain."));
        var selected = new HashSet<string>(assigned.Select(item => item is JsonObject ? item.Text("id") : JsonExtensions.Display(item)), StringComparer.Ordinal);
        var toggles = new List<CheckBox>();
        var all = new WrapPanel();
        var checkAll = new Button { Content = "Check all" }; checkAll.Click += (_, _) => { foreach (var check in toggles) check.IsChecked = true; }; all.Children.Add(checkAll);
        var uncheck = new Button { Content = "Uncheck all", Margin = new Thickness(8, 0, 0, 0) }; uncheck.Click += (_, _) => { foreach (var check in toggles) check.IsChecked = false; }; all.Children.Add(uncheck); content.Children.Add(all);
        foreach (var descriptor in session.Metadata!.Events)
        {
            var id = descriptor.Text("id"); var check = new CheckBox { Content = descriptor.Text("category") + " · " + descriptor.Text("label"), IsChecked = selected.Contains(id) };
            check.IsCheckedChanged += (_, _) => { if (check.IsChecked == true) selected.Add(id); else selected.Remove(id); }; content.Children.Add(check); toggles.Add(check);
        }
        await Dialogs.ChoiceAsync(this, "Assign “" + Path.GetFileName(path) + "”", content, [("Save to draft", "save"), ("Cancel", "cancel")], async choice =>
        {
            if (choice == "save") await session.MutateAsync("sound", new() { ["action"] = "assign", ["path"] = path, ["events"] = new JsonArray(selected.Select(id => (JsonNode?)JsonValue.Create(id)).ToArray()) });
        });
    }
    private async Task ShowConflictsAsync()
    {
        var result = await session.ReadAsync("command", new() { ["name"] = "shortcut-conflicts", ["arguments"] = new JsonObject() });
        await PresentConflictsAsync(result);
    }
}
