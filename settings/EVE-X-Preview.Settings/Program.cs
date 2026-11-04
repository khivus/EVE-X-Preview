using System.Text.Json;
using Avalonia;
using EveX.Settings.Bridge;

namespace EveX.Settings;

internal static class Program
{
    public static string Installation { get; private set; } = AppContext.BaseDirectory;
    public static SingleInstance? Instance { get; private set; }
    public static string? CaptureDirectory { get; private set; }

    [STAThread]
    public static int Main(string[] args)
    {
        try
        {
            if (args.Length == 2 && args[0] == "--probe")
            {
                var path = Path.GetFullPath(args[1]);
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                var temporary = path + "." + Environment.ProcessId + ".tmp";
                File.WriteAllText(temporary, JsonSerializer.Serialize(new { application = "EVE-X-Preview-Settings", bridgeVersion = 1, schemaVersion = 1 }));
                File.Move(temporary, path, true);
                return 0;
            }
            for (var i = 0; i < args.Length; i += 2)
            {
                if (i + 1 >= args.Length) throw new ArgumentException("An argument value is missing.");
                if (args[i] == "--installation") Installation = args[i + 1];
                else if (args[i] == "--capture-fixtures") CaptureDirectory = Path.GetFullPath(args[i + 1]);
                else throw new ArgumentException("Use --installation <directory>, --capture-fixtures <directory> or --probe <response-file>.");
            }
            Installation = InstallationIdentity.Normalize(Installation);
            Instance = new SingleInstance(Installation);
            if (!Instance.IsPrimary) { Instance.ActivateExisting(); Instance.Dispose(); return 0; }
            try { BuildAvaloniaApp().StartWithClassicDesktopLifetime(args); }
            finally { Instance.Dispose(); }
            return 0;
        }
        catch (Exception error)
        {
            // Startup failures must let AHK detect failure and use its internal editor.
            System.Diagnostics.Trace.WriteLine(error);
            return 1;
        }
    }

    public static AppBuilder BuildAvaloniaApp() => AppBuilder.Configure<App>().UsePlatformDetect().WithInterFont().LogToTrace();
}
