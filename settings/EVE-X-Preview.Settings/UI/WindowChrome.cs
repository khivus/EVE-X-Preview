using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using EveX.Settings.Bridge;
namespace EveX.Settings.UI;
public sealed partial class MainWindow {
    private Grid CreateHeader() {
        WindowDecorations=WindowDecorations.Full;ThemeController.AttachNativeFrame(this);
        var header=new Grid {ColumnDefinitions=new ColumnDefinitions("*,Auto"),Margin=new Thickness(16,5)};
        header.Children.Add(new TextBlock {Text="EVE-X-Preview",FontSize=16,FontWeight=FontWeight.SemiBold,VerticalAlignment=VerticalAlignment.Center});
        var row=new StackPanel {Orientation=Orientation.Horizontal,Spacing=10,VerticalAlignment=VerticalAlignment.Center};
        row.Children.Add(new TextBlock {Text="Profile",VerticalAlignment=VerticalAlignment.Center});
        profiles.Width=180;profiles.MinWidth=120;row.Children.Add(profiles);Grid.SetColumn(row,1);header.Children.Add(row);return header;
    }
}
internal static class NativeWindowFrame {
    private sealed class FrameState {public string? Signature;}
    private static readonly ConditionalWeakTable<Window,FrameState> states=new();
    [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(nint hwnd,int attribute,ref uint value,int size);
    [DllImport("dwmapi.dll")] private static extern int DwmGetWindowAttribute(nint hwnd,int attribute,out uint value,int size);
    public static void Apply(Window window,JsonObject? palette,bool force=false) {
        if(!OperatingSystem.IsWindows() || window.TryGetPlatformHandle()?.Handle is not {} hwnd || hwnd==0)return;
        var dark=palette is not null || window.ActualThemeVariant==ThemeVariant.Dark;
        uint ColorValue(string role) {var c=Color.Parse("#"+palette!.Text(role));return (uint)(c.R|(c.G<<8)|(c.B<<16));}
        uint caption=palette is null?0xFFFFFFFF:ColorValue("background"),text=palette is null?0xFFFFFFFF:ColorValue("foreground"),border=palette is null?0xFFFFFFFF:ColorValue("line");
        var signature=$"{dark}:{caption}:{text}:{border}";var state=states.GetOrCreateValue(window);if(!force&&state.Signature==signature)return;
        uint enabled=dark?1u:0u;if(DwmSetWindowAttribute(hwnd,20,ref enabled,4)!=0)DwmSetWindowAttribute(hwnd,19,ref enabled,4);
        // Exact palette colors are supported on Windows 11; Windows 10 retains its native dark caption.
        DwmSetWindowAttribute(hwnd,35,ref caption,4);DwmSetWindowAttribute(hwnd,36,ref text,4);DwmSetWindowAttribute(hwnd,34,ref border,4);state.Signature=signature;
    }
    [DllImport("user32.dll")] private static extern bool GetWindowRect(nint hwnd,out RECT rect);
    [DllImport("user32.dll",EntryPoint="SendMessageW")] private static extern nint SendMessage(nint hwnd,uint msg,nint wParam,nint lParam);
    [StructLayout(LayoutKind.Sequential)] private struct RECT {public int Left,Top,Right,Bottom;}
    internal static bool NativeCaptionHit(Window window) {
        var hwnd=window.TryGetPlatformHandle()!.Handle;if(!GetWindowRect(hwnd,out var rect))return false;
        var x=(rect.Left+rect.Right)/2;var y=rect.Top+20;
        return SendMessage(hwnd,0x84,0,(nint)((y<<16)|(x&0xffff)))==2;
    }
    public static bool DarkCaption(Window window)=>window.TryGetPlatformHandle()?.Handle is {} hwnd && DwmGetWindowAttribute(hwnd,20,out var v,4)==0&&v!=0;
}
