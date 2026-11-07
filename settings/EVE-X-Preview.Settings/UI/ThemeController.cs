using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Media;
using Avalonia.Styling;
using EveX.Settings.Bridge;

namespace EveX.Settings.UI;

// Palette values are exported by AHK. This maps presentation roles to Avalonia;
// adding an ordinary palette does not require another C# copy of its colors.
internal static class ThemeController
{
    private static string? applied;
    private static Styles? overrides;
    private static readonly Dictionary<string,SolidColorBrush> brushes=new();
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
        window.ActualThemeVariantChanged+=(_,_)=>NativeWindowFrame.Apply(window,nativePalette,true);
    }
    private static void RefreshNativeFrames(bool force=false) {
        if(Application.Current?.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
            foreach(var window in desktop.Windows)NativeWindowFrame.Apply(window,nativePalette,force);
    }

    public static IBrush TableSelection(bool row) => brushes[row?"EveTableRowSelectionBrush":"EveTableSelectionBrush"];
    public static IBrush Line => brushes["EveLineBrush"];

    public static void Apply(string mode, Metadata? metadata)
    {
        if(Application.Current is not {} app)return;
        var palette=metadata?.Value.Array("themes").OfType<System.Text.Json.Nodes.JsonObject>().FirstOrDefault(p=>p.Text("id")==mode);
        app.RequestedThemeVariant=palette is not null || mode=="dark"?ThemeVariant.Dark:mode=="light"?ThemeVariant.Light:ThemeVariant.Default;
        var dark=palette is not null || app.ActualThemeVariant==ThemeVariant.Dark;
        var signature=mode+"|"+palette?.ToJsonString()+"|"+dark;
        if(signature==applied)return;
        applied=signature;nativePalette=palette;
        Color Role(string key,string darkFallback,string lightFallback)=>Color.Parse("#"+(palette?.Text(key)??(dark?darkFallback:lightFallback)));
        SolidColorBrush Set(string key,Color color) {
            if(!brushes.TryGetValue(key,out var brush)) {brush=new SolidColorBrush(color);brushes[key]=brush;}
            else brush.Color=color;
            app.Resources[key]=brush;return brush;
        }
        Color Mix(Color a,Color b,double amount)=>Color.FromRgb((byte)Math.Round(a.R*(1-amount)+b.R*amount),(byte)Math.Round(a.G*(1-amount)+b.G*amount),(byte)Math.Round(a.B*(1-amount)+b.B*amount));
        var bg=Role("background","202124","F7F7F7");var ink=Role("foreground","ECEEF2","202020");
        var field=Role("input","2B2E33","FFFFFF");var band=Role("sidebar","191B1F","EFEFEF");
        var highlight=Role("accent","76B7FF","0078D4");var dim=Role("muted","A6ADB8","707070");
        var background=Set("EveBackgroundBrush",bg);var input=Set("EveInputBrush",field);
        var foreground=Set("EveForegroundBrush",ink);var muted=Set("EveMutedBrush",dim);
        var line=Set("EveLineBrush",Role("line","424750","DDE1E7"));var accent=Set("EveAccentBrush",highlight);
        var selection=Set("EveSelectionBrush",Role("selection","304C70","DCEBFA"));var chrome=Set("EveChromeBrush",band);
        var disabled=Set("EveDisabledBrush",Mix(field,band,.35));
        var selectionColor=((SolidColorBrush)selection).Color;Set("EveTableSelectionBrush",Color.FromArgb(65,selectionColor.R,selectionColor.G,selectionColor.B));Set("EveTableRowSelectionBrush",Color.FromArgb(90,selectionColor.R,selectionColor.G,selectionColor.B));
        var contrast=Set("EveAccentTextBrush",.299*highlight.R+.587*highlight.G+.114*highlight.B>145?Color.Parse("#202124"):Colors.White);
        void Resources(SolidColorBrush brush,params string[] keys) {foreach(var key in keys)app.Resources[key]=brush;}
        Resources(accent,"SystemControlHighlightAccentBrush","SystemControlForegroundAccentBrush","SliderTrackValueFill","SliderThumbBackground","SliderThumbBackgroundPointerOver","SliderThumbBackgroundPressed");
        Resources(foreground,"SystemControlForegroundBaseHighBrush","TextControlForeground","ComboBoxForeground","ButtonForeground","ToggleSwitchContentForeground");
        Resources(input,"ComboBoxDropDownBackground","TextControlBackground","TextControlBackgroundPointerOver","TextControlBackgroundFocused","ComboBoxBackground","ButtonBackground");
        Resources(selection,"ComboBoxItemBackgroundPointerOver","ComboBoxItemBackgroundSelected","ButtonBackgroundPointerOver","ButtonBackgroundPressed","ComboBoxBackgroundPointerOver","ComboBoxBackgroundPressed");
        Resources(disabled,"TextControlBackgroundDisabled","ComboBoxBackgroundDisabled","ButtonBackgroundDisabled");
        Resources(muted,"TextControlForegroundDisabled","TextControlPlaceholderForeground","TextControlPlaceholderForegroundDisabled","ButtonForegroundDisabled","ComboBoxForegroundDisabled","ToggleSwitchHeaderForegroundDisabled","RepeatButtonForegroundDisabled");
        Resources(line,"TextControlBorderBrush","TextControlBorderBrushPointerOver","ComboBoxBorderBrush","ComboBoxBorderBrushPointerOver","ButtonBorderBrush","ButtonBorderBrushPointerOver","SliderTrackFill");
        Resources(accent,"TextControlBorderBrushFocused");
        foreach(var state in new[]{"","PointerOver","Pressed"}) {
            Resources(accent,"ToggleSwitchFillOn"+state,"ToggleSwitchStrokeOn"+state);
            Resources(contrast,"ToggleSwitchKnobFillOn"+state);
            Resources(input,"ToggleSwitchFillOff"+state);Resources(muted,"ToggleSwitchStrokeOff"+state);
            Resources(foreground,"ToggleSwitchKnobFillOff"+state);
            app.Resources["ToggleSwitchContainerBackground"+state]=Brushes.Transparent;
        }
        Resources(disabled,"ToggleSwitchFillOnDisabled","ToggleSwitchFillOffDisabled");
        Resources(muted,"ToggleSwitchStrokeOnDisabled","ToggleSwitchStrokeOffDisabled","ToggleSwitchKnobFillOnDisabled","ToggleSwitchKnobFillOffDisabled");
        app.Resources["SystemAccentColor"]=highlight;
        if(overrides is null) {
            overrides=new Styles {
                new Style(s=>s.OfType<Window>()) {Setters={new Setter(TemplatedControl.BackgroundProperty,background),new Setter(TemplatedControl.ForegroundProperty,foreground)}},
                new Style(s=>s.OfType<Border>().Class("card")) {Setters={new Setter(Border.BackgroundProperty,input),new Setter(Border.BorderBrushProperty,line)}},
                new Style(s=>s.OfType<Border>().Class("sidebar")) {Setters={new Setter(Border.BackgroundProperty,chrome),new Setter(Border.BorderBrushProperty,line)}},
                new Style(s=>s.OfType<Border>().Class("chrome")) {Setters={new Setter(Border.BackgroundProperty,chrome),new Setter(Border.BorderBrushProperty,line)}},
                new Style(s=>s.OfType<Button>()) {Setters={new Setter(TemplatedControl.BackgroundProperty,input),new Setter(TemplatedControl.ForegroundProperty,foreground),new Setter(TemplatedControl.BorderBrushProperty,line)}},
                new Style(s=>s.OfType<Button>().Class("nav")) {Setters={new Setter(TemplatedControl.BackgroundProperty,Brushes.Transparent)}},
                new Style(s=>s.OfType<Button>().Class("nav").Class("selected")) {Setters={new Setter(TemplatedControl.BackgroundProperty,selection)}},
                new Style(s=>s.OfType<TextBlock>().Class("muted")) {Setters={new Setter(TextBlock.ForegroundProperty,muted),new Setter(Visual.OpacityProperty,1d)}},
                new Style(s=>s.OfType<TextBox>()) {Setters={new Setter(TemplatedControl.BackgroundProperty,background),new Setter(TemplatedControl.ForegroundProperty,foreground),new Setter(TemplatedControl.BorderBrushProperty,line),new Setter(TextBox.CaretBrushProperty,accent),new Setter(TextBox.SelectionBrushProperty,selection)}},
                new Style(s=>s.OfType<ComboBox>()) {Setters={new Setter(TemplatedControl.BackgroundProperty,background),new Setter(TemplatedControl.ForegroundProperty,foreground),new Setter(TemplatedControl.BorderBrushProperty,line)}}
            };app.Styles.Add(overrides);
        }
        RefreshNativeFrames(true);
    }
}
