using Avalonia.VisualTree;
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
    private Control CreateField(JsonObject field, bool horizontal = false)
    {
        var id = field.Text("id"); var label = field.Text("label");
        var state = session.Snapshot.Object("values").Object(id);
        var key = field.Bool("global") ? "global|" + sectionId + "|" + id : EditKey(id); var pending = session.GetPending(key);
        var current = pending?.Arguments["value"] ?? state["value"] ?? field["default"];
        var panel = new StackPanel { Spacing = 6 };
        var heading = SearchText(label + (field.Bool("global") && !horizontal ? "  ·  All profiles" : state.Bool("inherited") ? "  ·  From Default" : "")); heading.FontWeight = FontWeight.SemiBold;
        panel.Children.Add(heading);
        if (field.Text("help") is { Length: > 0 } help) panel.Children.Add(SearchText(help));
        void Change(JsonNode? value) => Stage(key, label, "set", new() { ["id"] = id, ["value"] = value?.DeepClone() });
        var input = horizontal && field.Text("kind")=="slider" ? VolumeInput(current,Change) : TypedInput(field, current, Change, horizontal);
        AutomationProperties.SetName(input, label);
        AutomationProperties.SetHelpText(input, field.Text("help"));
        input.IsEnabled = !state.Bool("inherited") && state.Bool("enabled", true);
        var feedback = new TextBlock { TextWrapping = TextWrapping.Wrap };
        BindFeedback(key, id, feedback, input);
        if(horizontal) {
            var controls=new StackPanel {Spacing=6};controls.Children.Add(input);
            if(field.Text("kind")=="slider")controls.Children.Add(Muted(field.Number("min")+" – "+field.Number("max")+" "+field.Text("unit")));
            controls.Children.Add(feedback);
            var row=new Grid {ColumnDefinitions=new ColumnDefinitions("*,240")};panel.Margin=new Thickness(0,0,18,0);
            row.Children.Add(panel);Grid.SetColumn(controls,1);row.Children.Add(controls);return row;
        }
        panel.Children.Add(input);panel.Children.Add(feedback);
        return panel;
    }

    private Control TypedInput(JsonObject field, JsonNode? current, Action<JsonNode?> changed, bool compact = false)
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
            var select = new ComboBox { ItemsSource = choices.Select(c => c.Text("label")).ToArray(), HorizontalAlignment = HorizontalAlignment.Left, MinWidth = compact ? 0 : 250, MaxWidth = 600 };
            select.SelectedIndex = Array.FindIndex(choices, choice => JsonExtensions.Display(choice["value"]) == value);
            select.SelectionChanged += (_, _) => { if (select.SelectedIndex >= 0) changed(choices[select.SelectedIndex]["value"]?.DeepClone()); }; return select;
        }
        if (kind == "hotkey") return new HotkeyEditor(value, text => changed(JsonValue.Create(text)), compact);
        if (kind == "font")
        {
            var fonts = session.Metadata!.Value.Array("fonts").Select(JsonExtensions.Display).Where(name=>!string.IsNullOrWhiteSpace(name)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            var selectedIndex=Array.FindIndex(fonts,name=>name.Equals(value.Trim(),StringComparison.OrdinalIgnoreCase));
            var font = new ComboBox { ItemsSource = fonts, SelectedIndex = selectedIndex, PlaceholderText=fonts.Length==0?"No installed fonts reported":selectedIndex<0?"Unavailable font: "+value:"Choose a font", HorizontalAlignment = HorizontalAlignment.Left, MinWidth = compact ? 0 : 250, MaxWidth = 600 };
            font.SelectionChanged += (_, _) => { if (font.SelectedItem is string selected) changed(JsonValue.Create(selected)); }; return font;
        }
        var text = new TextBox { Text = value, HorizontalAlignment = HorizontalAlignment.Left, MinWidth = 250, MaxWidth = kind == "multiline" ? double.PositiveInfinity : 600 };
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
            var numeric = new NumericUpDown { Width = compact ? double.NaN : 220, HorizontalAlignment = compact ? HorizontalAlignment.Stretch : HorizontalAlignment.Left,
                Minimum = decimal.MinValue, Maximum = decimal.MaxValue, Increment = kind == "number" ? .1m : 1m,
                FormatString = kind == "number" ? "0.###" : "0", ClipValueToMinMax = false,
                Value = decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var initial) ? initial : null };
            numeric.ValueChanged += (_, _) => changed(JsonValue.Create(numeric.Value?.ToString(CultureInfo.InvariantCulture) ?? ""));
            text.TextChanged += (_, _) => { if (decimal.TryParse(text.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var next) && numeric.Value != next) numeric.Value = next; };
            numeric.ValueChanged += (_, _) => { if (group.Children.OfType<Slider>().FirstOrDefault() is { } slider && numeric.Value is { } next && slider.Value != (double)next) slider.Value = (double)next; };
            numeric.AddHandler(TextBox.TextChangedEvent, (_, _) => {
                var editor = numeric.GetVisualDescendants().OfType<TextBox>().FirstOrDefault();
                if (editor is not null) changed(JsonValue.Create(editor.Text ?? ""));
            }, Avalonia.Interactivity.RoutingStrategies.Bubble);
            group.Children.Add(numeric); group.Children.Add(Muted(range)); return group;
        }
        if (kind is "directory" or "folder" or "file" or "sound")
        {
            text.MinWidth=0;text.MaxWidth=double.PositiveInfinity;text.HorizontalAlignment=HorizontalAlignment.Stretch;
            var row = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), MaxWidth = 600, HorizontalAlignment = HorizontalAlignment.Stretch };
            row.Children.Add(text);
            var browse = new Button { Content = "Browse", Margin = new Thickness(8, 0, 0, 0) }; Grid.SetColumn(browse, 1); row.Children.Add(browse);
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
            text.Width = double.NaN; text.MinWidth = 0; text.HorizontalAlignment=HorizontalAlignment.Stretch; text.Padding = new Thickness(compact ? 3 : 8, 4);
            text.FontSize = compact ? 12 : 14; text.PlaceholderText = "RRGGBB";
            var choose = new Button { Width = compact ? 30 : 36, Height = 34, MinHeight = 0, Padding = new Thickness(0), Margin = new Thickness(4,0,0,0), BorderThickness = new Thickness(1.5), Cursor = new Cursor(StandardCursorType.Hand) };
            choose.Classes.Add("color-swatch");
            var glyph = new TextBlock { Text = "✎", FontSize = 16, FontWeight = FontWeight.Bold };
            glyph.Classes.Add("color-swatch-icon");
            var iconBackground=new Border {Child=glyph,Padding=new Thickness(2,0),CornerRadius=new CornerRadius(3)};iconBackground.Classes.Add("color-swatch-icon-background");
            choose.HorizontalContentAlignment=HorizontalAlignment.Center;choose.VerticalContentAlignment=VerticalAlignment.Center;
            choose.Content = iconBackground; ToolTip.SetTip(choose, "Click the color preview to choose a color");
            AutomationProperties.SetName(choose, "Choose color");
            void UpdateSwatch()
            {
                var candidate = (text.Text ?? "").Replace("0x", "", StringComparison.OrdinalIgnoreCase).TrimStart('#');
                var valid = Color.TryParse("#" + candidate, out var color);
                choose.Background = valid ? new SolidColorBrush(color) : Brushes.Transparent;
            }
            UpdateSwatch(); text.TextChanged += (_, _) => UpdateSwatch();
            choose.Click += async (_, _) => { if (await PickColorAsync(text.Text ?? "") is { } selected) text.Text = selected; };
            var row = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), Width = compact ? double.NaN : 140, HorizontalAlignment = compact ? HorizontalAlignment.Stretch : HorizontalAlignment.Left };
            row.Children.Add(text); Grid.SetColumn(choose,1); row.Children.Add(choose); return row;
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
        var brightness = new Slider { Minimum = 0, Maximum = 100, Value = Math.Max(chosen.R, Math.Max(chosen.G, chosen.B)) * 100d / 255, TickFrequency = 1 };
        var brightnessModel = new EveX.Settings.State.ColorBrightness(); brightnessModel.SetRgb(chosen.R,chosen.G,chosen.B);
        var synchronizing = false;
        var valueLabel = new TextBlock();
        var output = new TextBlock { Text = $"{chosen.R:X2}{chosen.G:X2}{chosen.B:X2}", FontWeight = FontWeight.SemiBold };
        void Update()
        {
            if (synchronizing) return;
            chosen = Color.FromRgb((byte)Math.Round(red.Value), (byte)Math.Round(green.Value), (byte)Math.Round(blue.Value));
            brightnessModel.SetRgb(chosen.R,chosen.G,chosen.B);
            synchronizing=true; brightness.Value=brightnessModel.Percent; synchronizing=false;
            valueLabel.Text = "Brightness / value: " + Math.Round(brightness.Value) + "%";
            preview.Background = new SolidColorBrush(chosen); output.Text = $"{chosen.R:X2}{chosen.G:X2}{chosen.B:X2}";
        }
        foreach (var (label, slider) in new[] { ("Red", red), ("Green", green), ("Blue", blue) })
        {
            content.Children.Add(new TextBlock { Text = label }); content.Children.Add(slider);
            slider.PropertyChanged += (_, args) => { if (args.Property == Slider.ValueProperty) Update(); };
        }
        valueLabel.Text = "Brightness / value: " + Math.Round(brightness.Value) + "%";
        content.Children.Add(valueLabel); content.Children.Add(brightness);
        brightness.PropertyChanged += (_, args) => {
            if (args.Property != Slider.ValueProperty || synchronizing) return;
            var rgb = brightnessModel.WithValue(brightness.Value);
            synchronizing=true; red.Value=rgb.r; green.Value=rgb.g; blue.Value=rgb.b; synchronizing=false;
            chosen=Color.FromRgb(rgb.r,rgb.g,rgb.b); preview.Background=new SolidColorBrush(chosen); output.Text=$"{chosen.R:X2}{chosen.G:X2}{chosen.B:X2}";
            valueLabel.Text="Brightness / value: " + Math.Round(brightness.Value) + "%";
        };
        content.Children.Add(output);
        return await Dialogs.ChoiceAsync(this, "Choose RGB color", content, [("Use color", "use"), ("Cancel", "cancel")]) == "use" ? output.Text : null;
    }
}

internal sealed class HotkeyEditor : StackPanel
{
    private readonly bool compact;
    private readonly TextBox text;
    private readonly Button record;
    private bool recording;
    private TopLevel? top;
    public HotkeyEditor(string value, Action<string> changed, bool compact = false)
    {
        this.compact=compact;
        Spacing = 6;
        var row = new Grid { ColumnDefinitions = new ColumnDefinitions(compact ? "*,Auto,Auto,Auto" : "220,Auto,Auto,Auto"), HorizontalAlignment = compact ? HorizontalAlignment.Stretch : HorizontalAlignment.Left };
        text = new TextBox { Text = value, PlaceholderText = "No shortcut", IsReadOnly = true };
        record = new Button { Content = "Record", Margin = new Thickness(8, 0, 0, 0) };
        var clear = new Button { Content = "Clear", Margin = new Thickness(8, 0, 0, 0) };
        row.Children.Add(text); Grid.SetColumn(record, 1); row.Children.Add(record); Grid.SetColumn(clear, 2); row.Children.Add(clear); Children.Add(row);
        var advanced = new Button { Content = "Advanced", Margin = new Thickness(8, 0, 0, 0) }; Grid.SetColumn(advanced, 3); row.Children.Add(advanced);
        var help = new TextBlock { Text = "Record a shortcut, or enable advanced syntax (^ Ctrl, ! Alt, + Shift, # Win). Esc cancels recording. Mouse buttons may also be recorded.", TextWrapping = TextWrapping.Wrap, Opacity = .65, FontSize = 12 };
        ToolTip.SetTip(this, help.Text);
        advanced.Click += (_, _) => { text.IsReadOnly = !text.IsReadOnly; advanced.Content = text.IsReadOnly ? "Advanced" : "Done"; text.Focus(); };
        if (compact) {
            record.Content = "●"; clear.Content = "×"; advanced.Content = "A";
            foreach(var button in new[]{record,clear,advanced}) { button.Padding = new Thickness(0); button.Margin = new Thickness(3,0,0,0); button.Width=28;button.MinWidth=28;button.Height=34;button.HorizontalContentAlignment=HorizontalAlignment.Center;button.VerticalContentAlignment=VerticalAlignment.Center; }
            ToolTip.SetTip(record,"Record shortcut"); ToolTip.SetTip(clear,"Clear shortcut"); ToolTip.SetTip(advanced,"Advanced AutoHotkey syntax");
        }
        var lastText = value;
        text.TextChanged += (_, _) => { var candidate = text.Text ?? ""; if (candidate != lastText) { lastText = candidate; changed(candidate); } };
        clear.Click += (_, _) => { Stop(); text.Text = ""; };
        record.Click += (_, _) => { if (recording) Stop(); else Start(); };
        DetachedFromVisualTree += (_, _) => Stop();
    }
    private void Start()
    {
        recording = true; record.Content = compact ? "■" : "Press a shortcut"; top = TopLevel.GetTopLevel(this);
        top?.AddHandler(InputElement.KeyDownEvent, Capture, Avalonia.Interactivity.RoutingStrategies.Tunnel, true);
        top?.AddHandler(InputElement.PointerPressedEvent, CapturePointer, Avalonia.Interactivity.RoutingStrategies.Tunnel, true);
        record.Focus();
    }
    private void Stop()
    {
        recording = false; record.Content = compact ? "●" : "Record";
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
            Key.OemPlus => "=", Key.OemMinus => "-", Key.OemComma => ",", Key.OemPeriod => ".", Key.OemQuestion => "/", Key.OemSemicolon => ";", Key.OemQuotes => "'", Key.OemOpenBrackets => "[", Key.OemCloseBrackets => "]", Key.OemPipe => "\\", Key.OemTilde => "`",
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
