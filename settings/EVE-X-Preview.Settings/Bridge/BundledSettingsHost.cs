using System.Security.Cryptography;

namespace EveX.Settings.Bridge;

internal static class BundledSettingsHost
{
    internal static string? Extract()
    {
        using var resource=typeof(BundledSettingsHost).Assembly.GetManifestResourceStream("EveX.Settings.SettingsHost.exe");
        if(resource is null)return null; // Source builds can still use the adjacent AHK host.
        using var content=new MemoryStream();resource.CopyTo(content);var bytes=content.ToArray();
        var hash=Convert.ToHexStringLower(SHA256.HashData(bytes));
        var directory=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"EVE-X-Preview","SettingsHost",hash);
        Directory.CreateDirectory(directory);var path=Path.Combine(directory,"EVE-X-Preview-SettingsHost.exe");
        if(File.Exists(path)) {
            using var cached=File.OpenRead(path);
            if(cached.Length==bytes.Length && Convert.ToHexStringLower(SHA256.HashData(cached))==hash)return path;
        }
        var temporary=path+"."+Guid.NewGuid().ToString("N")+".tmp";
        try {File.WriteAllBytes(temporary,bytes);File.Move(temporary,path,true);}
        finally {if(File.Exists(temporary))File.Delete(temporary);}
        return path;
    }
}
