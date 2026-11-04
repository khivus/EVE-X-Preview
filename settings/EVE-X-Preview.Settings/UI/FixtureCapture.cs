using Avalonia;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.Threading;
using EveX.Settings.Bridge;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    // Deterministic render captures for the isolated Windows Sandbox fixture.
    // The 150% files simulate render DPI; they do not claim live display scaling.
    private async Task CaptureFixturesAsync(string directory)
    {
        if (session.Dirty) throw new InvalidOperationException("Capture needs a clean isolated fixture draft.");
        Directory.CreateDirectory(directory);
        var sections = new[] { "appearance", "character-shortcuts", "character-groups", "events", "sounds", "profiles", "inheritance", "interface", "diagnostics", "about" };
        var themes = new[] { "light", "dark" }.Concat(session.Metadata!.Value.Array("themes").Select(p => p.Text("id"))).Distinct();
        foreach (var theme in themes)
        foreach (var size in new[] { new Size(1100, 800), new Size(850, 650) })
        foreach (var id in sections.Where(id => session.Metadata!.Sections.Any(s => s.Text("id") == id)))
        {
            captureTheme = theme;
            Width = size.Width; Height = size.Height; sectionId = id; pageScroll.Offset = default;
            RenderNavigation(); RenderPage();
            await Task.Delay(300); // Let canonical specialized reads and layout settle.
            if (session.Pending.Count != 0) throw new InvalidOperationException("Rendering “" + id + "” unexpectedly staged an edit. Fixture capture stopped before accepting or saving anything.");
            ThemeController.Apply(theme, session.Metadata);
            await Dispatcher.UIThread.InvokeAsync(() => { InvalidateMeasure(); UpdateLayout(); }, DispatcherPriority.Render);
            await Task.Delay(70);
            foreach (var scale in new[] { 1.0, 1.5 })
            {
                using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(Bounds.Width * scale), (int)Math.Ceiling(Bounds.Height * scale)), new Vector(96 * scale, 96 * scale));
                bitmap.Render(this);
                bitmap.Save(Path.Combine(directory, $"{id}-{theme}-{size.Width:0}x{size.Height:0}-render{scale * 100:0}.png"), PngBitmapEncoderOptions.Default);
            }
        }
        captureTheme = null;
        await session.EndAsync(false);
        closing = true; Close();
    }
}
