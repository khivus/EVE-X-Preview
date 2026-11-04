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
        if (id == "about") page.Children.Add(Muted("EVE-X-Preview · Created by gonzo83. Maintained by khivus. Contributions by mrmjstc and CJKondur.\nSettings are validated and saved by the shared AutoHotkey settings system."));
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
        if (id.Contains("shortcut", StringComparison.Ordinal) || id is "character-groups" or "app-groups" or "profiles")
            page.Children.Add(ActionButton("Check shortcut conflicts", ShowConflictsAsync, false));
    }
    private void RenderProfiles()
    {
        var panel = new StackPanel { Spacing = 14 };
        var active = session.Snapshot.Array("profiles").FirstOrDefault(profile => profile.Bool("active"));
        panel.Children.Add(Muted("Active profile: " + active.Text("name", session.Profile) + "\nSelecting a profile changes the draft you are editing. Activation takes effect on Apply."));
        var buttons = new WrapPanel();
        buttons.Children.Add(ActionButton("Create profile…", CreateProfileAsync));
        var rename = ActionButton("Rename…", async () =>
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
        foreach (var section in session.Metadata!.Sections.Where(s => s.Text("group") != ""))
        {
            var id = section.Text("id");
            var check = new CheckBox { Content = section.Text("page") + " · " + section.Text("title"), IsChecked = flags.Bool(id) };
            check.IsCheckedChanged += (_, _) => flags[id] = check.IsChecked == true; content.Children.Add(check);
        }
        await Dialogs.ChoiceAsync(this, "Create profile", content, [("Create", "create"), ("Cancel", "cancel")], async choice =>
        {
            if (choice == "create") await session.MutateAsync("profile", new() { ["action"] = "add", ["name"] = name.Text, ["copyCurrent"] = copy.IsChecked == true, ["inheritance"] = flags.DeepClone() });
        });
    }
    private void RenderInheritance()
    {
        var panel = new StackPanel { Spacing = 10 };
        panel.Children.Add(Muted(session.Profile == "Default" ? "Default supplies the values for other profiles. Configure the creation template below." : "Inherited sections follow Default. Customize restores the profile's stored values."));
        if (session.Profile != "Default")
        {
            foreach (var section in session.Metadata!.Sections.Where(s => s.Text("group") != ""))
            {
                var id = section.Text("id");
                var check = new CheckBox { Content = section.Text("page") + " · " + section.Text("title"), IsChecked = session.Snapshot.Object("inheritance").Bool(id) };
                check.IsCheckedChanged += async (_, _) => await RunAsync(async () => { await session.MutateAsync("inheritance", new() { ["section"] = id, ["enabled"] = check.IsChecked == true }); });
                panel.Children.Add(check);
            }
        }
        page.Children.Add(Card(panel));
        var template = new StackPanel { Spacing = 10 };
        template.Children.Add(new TextBlock { Text = "New profile inheritance template", FontWeight = FontWeight.SemiBold });
        var templateKey = EditKey("creation-template");
        var flags = (session.GetPending(templateKey)?.Arguments["flags"] as JsonObject ?? session.Snapshot.Object("creationInheritance")).Copy();
        foreach (var section in session.Metadata!.Sections.Where(s => s.Text("group") != ""))
        {
            var id = section.Text("id"); var check = new CheckBox { Content = section.Text("page") + " · " + section.Text("title"), IsChecked = flags.Bool(id) };
            check.IsCheckedChanged += (_, _) => { flags[id] = check.IsChecked == true; Stage(templateKey, "Creation inheritance template", "creation-inheritance", new() { ["flags"] = flags.DeepClone() }); }; template.Children.Add(check);
        }
        template.Children.Add(ActionButton("Accept creation template", () => session.AcceptPendingAsync(templateKey))); page.Children.Add(Card(template));
    }
    private async void LoadGroupEditor(StackPanel host, string kind, bool inherited, int generation)
    {
        try
        {
            host.Children.Add(Muted("Loading groups…"));
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
            buttons.Children.Add(ActionButton("New group…", async () => { await Dialogs.NameActionAsync(this, "Create group", "", async name => { await session.MutateAsync("group", new() { ["kind"] = kind, ["action"] = "add", ["name"] = name }); selections[selectionKey] = name; }); }));
            if (selected != "")
            {
                var rename = ActionButton("Rename…", async () => { await Dialogs.NameActionAsync(this, "Rename group", selected, async name => { await session.MutateAsync("group", new() { ["kind"] = kind, ["action"] = "rename", ["name"] = selected, ["newName"] = name }); selections[selectionKey] = name; }); }); rename.Margin = new Thickness(8, 0, 0, 0); buttons.Children.Add(rename);
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
    private void RenderEvents(int generation)
    {
        var descriptors = session.Metadata!.Events.ToArray();
        if (descriptors.Length == 0) return;
        var selectionKey = session.Profile + "|event";
        var selected = selections.GetValueOrDefault(selectionKey, descriptors[0].Text("id"));
        var selector = new ComboBox { ItemsSource = descriptors.Select(e => e.Text("category") + " · " + e.Text("label")).ToArray(), SelectedIndex = Math.Max(0, Array.FindIndex(descriptors, e => e.Text("id") == selected)), HorizontalAlignment = HorizontalAlignment.Stretch };
        selector.SelectionChanged += (_, _) => { if (selector.SelectedIndex >= 0) { selections[selectionKey] = descriptors[selector.SelectedIndex].Text("id"); RenderPage(); } }; page.Children.Add(selector);
        var host = new StackPanel { Spacing = 14 }; page.Children.Add(Card(host)); LoadEventEditor(host, descriptors[selector.SelectedIndex].Text("id"), generation);
    }
    private async void LoadEventEditor(StackPanel host, string id, int generation)
    {
        try
        {
            host.Children.Add(Muted("Loading event conditions…"));
            var data = await session.ReadAsync("event", new() { ["id"] = id });
            if (generation != pageGeneration) return;
            host.Children.Clear();
            var editKey = EditKey("event|" + id);
            var values = (session.GetPending(editKey)?.Arguments["values"] as JsonObject ?? data.Object("values")).Copy();
            var fields = data.Array("fields").OfType<JsonObject>().ToArray();
            JsonObject EditableValues()
            {
                var editable = new JsonObject();
                foreach (var field in fields) editable[field.Text("id")] = values[field.Text("id")]?.DeepClone();
                return editable;
            }
            foreach (var field in fields)
            {
                var fieldId = field.Text("id");
                host.Children.Add(new TextBlock { Text = field.Text("label"), FontWeight = FontWeight.SemiBold });
                if (field.Text("help") is { Length: > 0 } help) host.Children.Add(Muted(help));
                var control = TypedInput(field, values[fieldId], value => { values[fieldId] = value?.DeepClone(); Stage(editKey, data.Text("label"), "event", new() { ["id"] = id, ["values"] = EditableValues() }); });
                control.IsEnabled = !data.Bool("inherited"); host.Children.Add(control);
            }
            var actions = new WrapPanel();
            actions.Children.Add(ActionButton("Accept event edits", () => session.AcceptPendingAsync(editKey)));
            if (session.GetPending(editKey) is not null) actions.Children.Add(ActionButton("Discard unfinished event edits", () => session.DiscardPendingAsync(editKey)));
            if (session.Metadata!.Commands.FirstOrDefault(command => command.Text("id") == "preview-sound") is { } preview)
            {
                var button = ActionButton("Preview this event sound", () => RunCommandAsync(preview, new() { ["event"] = id, ["path"] = values["sound"]?.DeepClone(), ["volume"] = values["soundVolume"]?.DeepClone() }), false); button.Margin = new Thickness(8, 0, 0, 0); button.IsEnabled = !preview.Bool("runtimeOnly") || session.Role != "offline"; actions.Children.Add(button);
            }
            host.Children.Add(actions);
            if (session.GetPending(editKey)?.Error is { } error) host.Children.Add(new TextBlock { Text = error, Classes = { "error" } });
        }
        catch (Exception e) { if (generation == pageGeneration) { host.Children.Clear(); host.Children.Add(new TextBlock { Text = e.Message, Classes = { "error" } }); } }
    }
    private void RenderSounds(int generation)
    {
        var host = new StackPanel { Spacing = 14 }; page.Children.Add(Card(host)); LoadSounds(host, generation);
    }
    private async void LoadSounds(StackPanel host, int generation)
    {
        try
        {
            host.Children.Add(Muted("Loading sound library…"));
            var data = await session.ReadAsync("sound", new() { ["action"] = "list" });
            if (generation != pageGeneration) return;
            host.Children.Clear();
            host.Children.Add(ActionButton("Add sound files…", async () =>
            {
                await ProtectModalAsync(true);
                try
                {
                var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions { Title = "Add sounds to the shared library", AllowMultiple = true, FileTypeFilter = [new FilePickerFileType("Audio") { Patterns = ["*.wav", "*.mp3", "*.ogg", "*.flac", "*.m4a", "*.wma"] }, FilePickerFileTypes.All] });
                foreach (var file in files) if (file.TryGetLocalPath() is { } path) await session.MutateAsync("sound", new() { ["action"] = "add", ["path"] = path });
                }
                finally { await ProtectModalAsync(false); }
            }));
            host.Children.Add(Muted("Library and file gain are shared. Event assignments belong to the selected profile. Removing a library reference keeps the original file and event assignments; Undo restores it."));
            foreach (var sound in data.Array("rows").OfType<JsonObject>())
            {
                var path = sound.Text("path"); var key = EditKey("gain|" + path);
                var row = new StackPanel { Spacing = 8 };
                row.Children.Add(new TextBlock { Text = Path.GetFileName(path), FontWeight = FontWeight.SemiBold, FontSize = 16 }); row.Children.Add(Muted(path));
                var assigned = string.Join(", ", sound.Array("assigned").Select(item => item is JsonObject ? item.Text("label", item.Text("id")) : JsonExtensions.Display(item)));
                row.Children.Add(Muted(assigned == "" ? "No event assignments in this profile" : "Assigned: " + assigned));
                var gain = session.GetPending(key)?.Arguments["volume"] ?? sound["volume"];
                row.Children.Add(new TextBlock { Text = "File gain (%)" });
                var gainInput = TypedInput(new() { ["kind"] = "slider", ["min"] = 0, ["max"] = 300, ["unit"] = "%" }, gain, value => Stage(key, "File gain: " + Path.GetFileName(path), "sound", new() { ["action"] = "gain", ["path"] = path, ["volume"] = value?.DeepClone() })); row.Children.Add(gainInput);
                var actions = new WrapPanel();
                actions.Children.Add(ActionButton("Accept file gain", () => session.AcceptPendingAsync(key)));
                if (session.GetPending(key) is not null) actions.Children.Add(ActionButton("Discard unfinished gain", () => session.DiscardPendingAsync(key)));
                var assign = ActionButton("Assign to events…", () => AssignSoundAsync(path, sound.Array("assigned"))); assign.Margin = new Thickness(8, 0, 0, 0); actions.Children.Add(assign);
                var remove = ActionButton("Remove reference", async () => { if (await Dialogs.ConfirmAsync(this, "Remove library reference?", "Keep the original file and event assignments, and remove this library reference?")) await session.MutateAsync("sound", new() { ["action"] = "remove", ["path"] = path }); }); remove.Margin = new Thickness(8, 0, 0, 0); actions.Children.Add(remove);
                if (session.Metadata!.Commands.FirstOrDefault(command => command.Text("id") == "preview-sound") is { } preview)
                {
                    var play = ActionButton("Preview", async () => { await session.AcceptPendingAsync(key); await RunCommandAsync(preview, new() { ["path"] = path }); }, false); play.IsEnabled = !preview.Bool("runtimeOnly") || session.Role != "offline"; play.Margin = new Thickness(8, 0, 0, 0); actions.Children.Add(play);
                }
                row.Children.Add(actions); host.Children.Add(new Border { Child = row, Padding = new Thickness(0, 14, 0, 20), BorderThickness = new Thickness(0, 0, 0, 1), BorderBrush = Brushes.Gray });
            }
        }
        catch (Exception e) { if (generation == pageGeneration) { host.Children.Clear(); host.Children.Add(new TextBlock { Text = e.Message, Classes = { "error" } }); } }
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
