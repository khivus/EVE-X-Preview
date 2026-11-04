using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Automation;
using EveX.Settings.Bridge;
using System.Globalization;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    private Control CreateField(JsonObject field)
    {
        var id = field.Text("id"); var label = field.Text("label");
        var state = session.Snapshot.Object("values").Object(id);
        var key = field.Bool("global") ? "global|" + sectionId + "|" + id : EditKey(id); var pending = session.GetPending(key);
        var current = pending?.Arguments["value"] ?? state["value"] ?? field["default"];
        var panel = new StackPanel { Spacing = 6 };
        var heading = new TextBlock { Text = label + (field.Bool("global") ? "  ·  All profiles" : state.Bool("inherited") ? "  ·  From Default" : ""), FontWeight = FontWeight.SemiBold, TextWrapping = TextWrapping.Wrap };
        panel.Children.Add(heading);
        if (field.Text("help") is { Length: > 0 } help) panel.Children.Add(Muted(help));
        var input = TypedInput(field, current, value => Stage(key, label, "set", new() { ["id"] = id, ["value"] = value?.DeepClone() }));
        AutomationProperties.SetName(input, label);
        AutomationProperties.SetHelpText(input, field.Text("help"));
        input.IsEnabled = !state.Bool("inherited") && state.Bool("enabled", true);
        panel.Children.Add(input);
        if (pending?.Error is { } message) panel.Children.Add(new TextBlock { Text = message, Classes = { "error" } });
        if (pending is not null)
        {
            var acceptField = ActionButton("Accept this edit", () => session.AcceptPendingAsync(key));
            acceptField.HorizontalAlignment = HorizontalAlignment.Left; panel.Children.Add(acceptField);
            panel.Children.Add(ActionButton("Discard this unfinished edit", () => session.DiscardPendingAsync(key)));
        }
        return panel;
    }

    private Control TypedInput(JsonObject field, JsonNode? current, Action<JsonNode?> changed)
    {
        var kind = field.Text("kind"); var value = JsonExtensions.Display(current);
        // Avalonia can deliver the initial TextChanged/SelectionChanged after
        // attachment. Treat equivalent initial values as presentation setup,
        // including AHK's numeric booleans and numeric values shown as text.
        var notify = changed;
        var lastValue = kind is "bool" or "boolean" ? (JsonExtensions.IsTrue(current) ? "true" : "false") : value;
        changed = candidate =>
        {
            var candidateValue = kind is "bool" or "boolean" ? (JsonExtensions.IsTrue(candidate) ? "true" : "false") : JsonExtensions.Display(candidate);
            if (candidateValue == lastValue) return;
            lastValue = candidateValue; notify(candidate);
        };
        if (kind is "bool" or "boolean")
        {
            var toggle = new ToggleSwitch { IsChecked = JsonExtensions.IsTrue(current), OnContent = "On", OffContent = "Off" };
            toggle.IsCheckedChanged += (_, _) => changed(JsonValue.Create(toggle.IsChecked == true)); return toggle;
        }
        var choices = (kind == "gesture" ? session.Metadata!.Value.Array("gestures") : field.Array("choices")).OfType<JsonObject>().ToArray();
        if (choices.Length != 0)
        {
            var select = new ComboBox { ItemsSource = choices.Select(c => c.Text("label")).ToArray(), HorizontalAlignment = HorizontalAlignment.Stretch, MaxWidth = 600 };
            select.SelectedIndex = Array.FindIndex(choices, choice => JsonExtensions.Display(choice["value"]) == value);
            select.SelectionChanged += (_, _) => { if (select.SelectedIndex >= 0) changed(choices[select.SelectedIndex]["value"]?.DeepClone()); }; return select;
        }
        if (kind == "hotkey") return new HotkeyEditor(value, text => changed(JsonValue.Create(text)));
        if (kind == "font")
        {
            var fonts = session.Metadata!.Value.Array("fonts").Select(JsonExtensions.Display).ToList();
            if (fonts.Count == 0) fonts = FontManager.Current.SystemFonts.Select(font => font.Name).Distinct().OrderBy(name => name).ToList();
            if (!fonts.Contains(value)) fonts.Insert(0, value);
            var font = new ComboBox { ItemsSource = fonts, SelectedItem = value, HorizontalAlignment = HorizontalAlignment.Stretch, MaxWidth = 600 };
            font.SelectionChanged += (_, _) => { if (font.SelectedItem is string selected) changed(JsonValue.Create(selected)); }; return font;
        }
        var text = new TextBox { Text = value, HorizontalAlignment = HorizontalAlignment.Stretch, MaxWidth = kind == "multiline" ? double.PositiveInfinity : 600 };
        text.TextChanged += (_, _) => changed(JsonValue.Create(text.Text ?? ""));
        if (kind == "readonly") { text.IsReadOnly = true; return text; }
        if (kind == "multiline") { text.AcceptsReturn = true; text.TextWrapping = TextWrapping.Wrap; text.MinHeight = 120; text.MaxHeight = 320; return text; }
        if (kind is "integer" or "number" or "slider")
        {
            var group = new StackPanel { Spacing = 6 };
            var range = field.Number("min").ToString(CultureInfo.InvariantCulture) + " – " + field.Number("max").ToString(CultureInfo.InvariantCulture) + " " + field.Text("unit");
            text.PlaceholderText = range; text.MaxWidth = 220; text.HorizontalAlignment = HorizontalAlignment.Left;
            if (kind == "slider")
            {
                var slider = new Slider { Minimum = field.Number("min"), Maximum = Math.Max(field.Number("min") + 1, field.Number("max")), Value = double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) ? parsed : field.Number("min"), TickFrequency = 5, MaxWidth = 600, HorizontalAlignment = HorizontalAlignment.Stretch };
                slider.PropertyChanged += (_, args) => { if (args.Property == Slider.ValueProperty) text.Text = Math.Round(slider.Value).ToString(CultureInfo.InvariantCulture); };
                group.Children.Add(slider);
            }
            group.Children.Add(text); group.Children.Add(Muted(range)); return group;
        }
        if (kind is "directory" or "folder" or "file" or "sound")
        {
            var row = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), MaxWidth = 700, HorizontalAlignment = HorizontalAlignment.Stretch };
            row.Children.Add(text);
            var browse = new Button { Content = "Browse…", Margin = new Thickness(8, 0, 0, 0) }; Grid.SetColumn(browse, 1); row.Children.Add(browse);
            browse.Click += async (_, _) =>
            {
                await ProtectModalAsync(true);
                try
                {
                if (kind is "directory" or "folder")
                {
                    var selection = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions { Title = field.Text("label"), AllowMultiple = false });
                    if (selection.FirstOrDefault()?.TryGetLocalPath() is { } path) text.Text = path;
                }
                else
                {
                    var selection = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions { Title = field.Text("label"), AllowMultiple = false });
                    if (selection.FirstOrDefault()?.TryGetLocalPath() is { } path) text.Text = path;
                }
                }
                catch (Exception error) { ShowError(error); }
                finally { await ProtectModalAsync(false); }
            };
            return row;
        }
        if (kind == "color")
        {
            text.PlaceholderText = "RGB hex, for example FF9900";
            var swatch = new Border { Width = 34, Height = 34, CornerRadius = new CornerRadius(6), BorderBrush = Brushes.Gray, BorderThickness = new Thickness(1), Margin = new Thickness(8, 0, 0, 0) };
            void UpdateSwatch()
            {
                var candidate = (text.Text ?? "").Replace("0x", "", StringComparison.OrdinalIgnoreCase).TrimStart('#');
                swatch.Background = Color.TryParse("#" + candidate, out var color) ? new SolidColorBrush(color) : Brushes.Transparent;
            }
            UpdateSwatch(); text.TextChanged += (_, _) => UpdateSwatch();
            var choose = new Button { Content = "Choose…", Margin = new Thickness(8, 0, 0, 0) };
            choose.Click += async (_, _) => { if (await PickColorAsync(text.Text ?? "") is { } selected) text.Text = selected; };
            var row = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto"), MaxWidth = 700 }; row.Children.Add(text); Grid.SetColumn(swatch, 1); row.Children.Add(swatch); Grid.SetColumn(choose, 2); row.Children.Add(choose); return row;
        }
        return text;
    }
    private async Task<string?> PickColorAsync(string initial)
    {
        var cleaned = initial.Replace("0x", "", StringComparison.OrdinalIgnoreCase).TrimStart('#');
        var chosen = Color.TryParse("#" + cleaned, out var existing) ? existing : Colors.White;
        var content = new StackPanel { Spacing = 14 };
        var preview = new Border { Height = 65, CornerRadius = new CornerRadius(8), Background = new SolidColorBrush(chosen), BorderBrush = Brushes.Gray, BorderThickness = new Thickness(1) }; content.Children.Add(preview);
        var red = new Slider { Minimum = 0, Maximum = 255, Value = chosen.R, TickFrequency = 1 };
        var green = new Slider { Minimum = 0, Maximum = 255, Value = chosen.G, TickFrequency = 1 };
        var blue = new Slider { Minimum = 0, Maximum = 255, Value = chosen.B, TickFrequency = 1 };
        var output = new TextBlock { Text = $"{chosen.R:X2}{chosen.G:X2}{chosen.B:X2}", FontWeight = FontWeight.SemiBold };
        void Update()
        {
            chosen = Color.FromRgb((byte)Math.Round(red.Value), (byte)Math.Round(green.Value), (byte)Math.Round(blue.Value));
            preview.Background = new SolidColorBrush(chosen); output.Text = $"{chosen.R:X2}{chosen.G:X2}{chosen.B:X2}";
        }
        foreach (var (label, slider) in new[] { ("Red", red), ("Green", green), ("Blue", blue) })
        {
            content.Children.Add(new TextBlock { Text = label }); content.Children.Add(slider);
            slider.PropertyChanged += (_, args) => { if (args.Property == Slider.ValueProperty) Update(); };
        }
        content.Children.Add(output);
        return await Dialogs.ChoiceAsync(this, "Choose RGB color", content, [("Use color", "use"), ("Cancel", "cancel")]) == "use" ? output.Text : null;
    }
}

internal sealed class HotkeyEditor : StackPanel
{
    private readonly TextBox text;
    private readonly Button record;
    private bool recording;
    private TopLevel? top;
    public HotkeyEditor(string value, Action<string> changed)
    {
        Spacing = 6;
        var row = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto,Auto"), MaxWidth = 700, HorizontalAlignment = HorizontalAlignment.Stretch };
        text = new TextBox { Text = value, PlaceholderText = "No shortcut", IsReadOnly = true };
        record = new Button { Content = "Record", Margin = new Thickness(8, 0, 0, 0) };
        var clear = new Button { Content = "Clear", Margin = new Thickness(8, 0, 0, 0) };
        row.Children.Add(text); Grid.SetColumn(record, 1); row.Children.Add(record); Grid.SetColumn(clear, 2); row.Children.Add(clear); Children.Add(row);
        var advanced = new CheckBox { Content = "Advanced AutoHotkey syntax" }; Children.Add(advanced);
        var help = new TextBlock { Text = "Record a shortcut, or enable advanced syntax (^ Ctrl, ! Alt, + Shift, # Win). Esc cancels recording. Mouse buttons may also be recorded.", TextWrapping = TextWrapping.Wrap, Opacity = .65, FontSize = 12 };
        Children.Add(help);
        advanced.IsCheckedChanged += (_, _) => text.IsReadOnly = advanced.IsChecked != true;
        var lastText = value;
        text.TextChanged += (_, _) => { var candidate = text.Text ?? ""; if (candidate != lastText) { lastText = candidate; changed(candidate); } };
        clear.Click += (_, _) => { Stop(); text.Text = ""; };
        record.Click += (_, _) => { if (recording) Stop(); else Start(); };
        DetachedFromVisualTree += (_, _) => Stop();
    }
    private void Start()
    {
        recording = true; record.Content = "Press a shortcut…"; top = TopLevel.GetTopLevel(this);
        top?.AddHandler(InputElement.KeyDownEvent, Capture, Avalonia.Interactivity.RoutingStrategies.Tunnel, true);
        top?.AddHandler(InputElement.PointerPressedEvent, CapturePointer, Avalonia.Interactivity.RoutingStrategies.Tunnel, true);
        record.Focus();
    }
    private void Stop()
    {
        recording = false; record.Content = "Record";
        top?.RemoveHandler(InputElement.KeyDownEvent, Capture); top?.RemoveHandler(InputElement.PointerPressedEvent, CapturePointer); top = null;
    }
    private void Capture(object? sender, KeyEventArgs args)
    {
        if (!recording) return;
        args.Handled = true;
        if (args.Key == Key.Escape) { Stop(); return; }
        if (args.Key is Key.LeftAlt or Key.RightAlt or Key.LeftCtrl or Key.RightCtrl or Key.LeftShift or Key.RightShift or Key.LWin or Key.RWin) return;
        var key = args.Key switch
        {
            Key.Return => "Enter", Key.Next => "PgDn", Key.Prior => "PgUp", Key.Back => "Backspace",
            Key.OemPlus => "=", Key.OemMinus => "-", Key.OemComma => ",", Key.OemPeriod => ".", Key.OemQuestion => "/", Key.OemSemicolon => ";", Key.OemQuotes => "'", Key.OemOpenBrackets => "[", Key.OemCloseBrackets => "]", Key.OemPipe => "\\", Key.OemTilde => "``",
            >= Key.D0 and <= Key.D9 => ((int)args.Key - (int)Key.D0).ToString(),
            _ => args.Key.ToString()
        };
        text.Text = Prefix(args.KeyModifiers) + key; Stop();
    }
    private void CapturePointer(object? sender, PointerPressedEventArgs args)
    {
        if (!recording || args.Source == record) return;
        var point = args.GetCurrentPoint(this);
        var key = point.Properties.IsXButton1Pressed ? "XButton1" : point.Properties.IsXButton2Pressed ? "XButton2" : point.Properties.IsMiddleButtonPressed ? "MButton" : point.Properties.IsRightButtonPressed ? "RButton" : "LButton";
        args.Handled = true; text.Text = Prefix(args.KeyModifiers) + key; Stop();
    }
    private static string Prefix(KeyModifiers modifiers) => (modifiers.HasFlag(KeyModifiers.Control) ? "^" : "") + (modifiers.HasFlag(KeyModifiers.Alt) ? "!" : "") + (modifiers.HasFlag(KeyModifiers.Shift) ? "+" : "") + (modifiers.HasFlag(KeyModifiers.Meta) ? "#" : "");
}
