using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Themes.Fluent;
using EveX.Settings.Bridge;

namespace EveX.Settings.UI;

// Palette values are exported by AHK. This maps presentation roles to Avalonia;
// adding an ordinary palette does not require another C# copy of its colors.
internal static class ThemeController
{
    private static string? applied;
    private static Styles? overrides;
    private static FluentTheme? fluent;
    private static ColorPaletteResources? originalDark;
    private static bool capturedOriginal;
    private static System.Text.Json.Nodes.JsonObject? nativePalette;
    // Presentation-only cache: the canonical configuration is still owned by AHK.
    private static string ThemeCachePath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"EVE-X-Preview","SettingsEditor",InstallationIdentity.Hash(Program.Installation)+".theme.json");
    public static void RestoreStartupTheme(string? cachePath=null) {
        if(cachePath is not null || Program.CaptureDirectory is null) {
            try {
                var file=new FileInfo(cachePath??ThemeCachePath);
                if(file.Exists && file.Length<=8192 && System.Text.Json.Nodes.JsonNode.Parse(File.ReadAllText(file.FullName)) is System.Text.Json.Nodes.JsonObject saved) {
                    var palette=saved["palette"] as System.Text.Json.Nodes.JsonObject;
                    if(palette is not null)foreach(var key in new[]{"background","sidebar","input","foreground","muted","line","accent","selection"})if(!Color.TryParse("#"+palette.Text(key),out _))throw new InvalidDataException("Invalid cached palette");
                    Apply(saved.Text("mode","dark"),new Metadata(new(){["themes"]=new System.Text.Json.Nodes.JsonArray(palette?.DeepClone())}));return;
                }
            } catch { /* An unavailable cache must not prevent opening settings. */ }
        }
        Apply("dark",null);
    }
    public static void RememberTheme(string mode,Metadata metadata,string? cachePath=null) {
        if(cachePath is null && Program.CaptureDirectory is not null)return;
        try {
            var palette=metadata.Value.Array("themes").FirstOrDefault(p=>p.Text("id")==mode);
            var saved=new System.Text.Json.Nodes.JsonObject {["mode"]=mode,["palette"]=palette?.DeepClone()};
            var path=cachePath??ThemeCachePath;Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            var temporary=path+"."+Environment.ProcessId+".tmp";File.WriteAllText(temporary,saved.ToJsonString());File.Move(temporary,path,true);
        } catch { /* Theme caching is optional, independent of canonical saving. */ }
    }
    public static void AttachNativeFrame(Window window) {
        NativeWindowFrame.Apply(window,nativePalette);
        window.Initialized+=(_,_)=>NativeWindowFrame.Apply(window,nativePalette,true);
        window.Opened+=(_,_)=>NativeWindowFrame.Apply(window,nativePalette,true);
        window.ActualThemeVariantChanged+=(_,_)=>NativeWindowFrame.Apply(window,nativePalette);
    }
    private static void RefreshNativeFrames() {
        if(Application.Current?.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
            foreach(var window in desktop.Windows)NativeWindowFrame.Apply(window,nativePalette);
    }

    public static void Apply(string mode, Metadata? metadata)
    {
        if (Application.Current is not { } app) return;
        var palette = metadata?.Value.Array("themes").OfType<System.Text.Json.Nodes.JsonObject>().FirstOrDefault(p => p.Text("id") == mode);
        var signature = mode + "|" + palette?.ToJsonString();
        if (signature == applied) {RefreshNativeFrames();return;}
        applied = signature; nativePalette=palette;
        if (overrides is not null) { app.Styles.Remove(overrides); overrides = null; }
        fluent ??= app.Styles.OfType<FluentTheme>().FirstOrDefault();
        if (fluent is not null)
        {
            if (!capturedOriginal) { fluent.Palettes.TryGetValue(ThemeVariant.Dark, out originalDark); capturedOriginal = true; }
            if (originalDark is not null) fluent.Palettes[ThemeVariant.Dark] = originalDark;
            else fluent.Palettes.Remove(ThemeVariant.Dark);
        }
        app.RequestedThemeVariant = palette is not null || mode == "dark" ? ThemeVariant.Dark : mode == "light" ? ThemeVariant.Light : ThemeVariant.Default;
        if (palette is null) {RefreshNativeFrames();return;}

        Color ColorRole(string key) => Color.Parse("#" + palette.Text(key));
        SolidColorBrush Brush(string key) => new(ColorRole(key));
        var background = Brush("background"); var input = Brush("input");
        var foreground = Brush("foreground"); var muted = Brush("muted");
        var line = Brush("line"); var accent = Brush("accent"); var selection = Brush("selection");
        if (fluent is not null)
            fluent.Palettes[ThemeVariant.Dark] = new ColorPaletteResources
            {
                Accent = ColorRole("accent"), AltHigh = ColorRole("background"), AltLow = ColorRole("background"),
                AltMedium = ColorRole("input"), AltMediumHigh = ColorRole("input"), AltMediumLow = ColorRole("sidebar"),
                BaseHigh = ColorRole("foreground"), BaseMediumHigh = ColorRole("foreground"),
                BaseMedium = ColorRole("muted"), BaseMediumLow = ColorRole("muted"), BaseLow = ColorRole("line"),
                ChromeAltLow = ColorRole("sidebar"), ChromeHigh = ColorRole("line"), ChromeMedium = ColorRole("input"),
                ChromeMediumLow = ColorRole("sidebar"), ChromeLow = ColorRole("background"),
                ChromeBlackHigh = ColorRole("foreground"), ChromeWhite = ColorRole("foreground"),
                ListLow = ColorRole("input"), ListMedium = ColorRole("selection"), RegionColor = ColorRole("background")
            };
        overrides = new Styles
        {
            new Style(s => s.OfType<Button>().Class("color-swatch")) { Setters = { new Setter(TemplatedControl.BorderBrushProperty, accent) } },
            new Style(s => s.OfType<TextBlock>().Class("color-swatch-icon")) { Setters = { new Setter(TextBlock.ForegroundProperty, foreground) } },
            new Style(s => s.OfType<Border>().Class("color-swatch-icon-background")) { Setters = { new Setter(Border.BackgroundProperty, input) } },
            new Style(s => s.OfType<Window>()) { Setters = { new Setter(TemplatedControl.BackgroundProperty, background), new Setter(TemplatedControl.ForegroundProperty, foreground) } },
            new Style(s => s.OfType<Border>().Class("card")) { Setters = { new Setter(Border.BackgroundProperty, input), new Setter(Border.BorderBrushProperty, line) } },
            new Style(s => s.OfType<Border>().Class("sidebar")) { Setters = { new Setter(Border.BackgroundProperty, Brush("sidebar")), new Setter(Border.BorderBrushProperty, line) } },
            new Style(s => s.OfType<Button>()) { Setters = { new Setter(TemplatedControl.BackgroundProperty, input), new Setter(TemplatedControl.ForegroundProperty, foreground), new Setter(TemplatedControl.BorderBrushProperty, line) } },
            new Style(s => s.OfType<Button>().Class("nav")) { Setters = { new Setter(TemplatedControl.BackgroundProperty, Brushes.Transparent) } },
            new Style(s => s.OfType<Button>().Class("nav").Class("selected")) { Setters = { new Setter(TemplatedControl.BackgroundProperty, selection) } },
            new Style(s => s.OfType<TextBlock>().Class("muted")) { Setters = { new Setter(TextBlock.ForegroundProperty, muted), new Setter(Visual.OpacityProperty, 1d) } },
            new Style(s => s.OfType<TextBox>()) { Setters = { new Setter(TemplatedControl.BackgroundProperty, background), new Setter(TemplatedControl.ForegroundProperty, foreground), new Setter(TemplatedControl.BorderBrushProperty, line), new Setter(TextBox.CaretBrushProperty, accent), new Setter(TextBox.SelectionBrushProperty, selection) } },
            new Style(s => s.OfType<ComboBox>()) { Setters = { new Setter(TemplatedControl.BackgroundProperty, background), new Setter(TemplatedControl.ForegroundProperty, foreground), new Setter(TemplatedControl.BorderBrushProperty, line) } }
        };
        app.Styles.Add(overrides);RefreshNativeFrames();
    }
}
