using System.ComponentModel;
using System.Runtime.InteropServices;

namespace EveX.Settings.Bridge;

// Own a real top-level HWND on its own message thread. A message-only HWND
// would not be discoverable by the AHK endpoint's sender checks.
internal sealed class NativeWindow : IDisposable
{
    private readonly Thread thread;
    private readonly TaskCompletionSource<nint> ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly Native.WindowProcedure procedure;
    private string className = "";
    public nint Handle => ready.Task.GetAwaiter().GetResult();
    public NativeWindow()
    {
        procedure = WindowProcedure;
        thread = new Thread(Run) { IsBackground = true, Name = "Settings bridge sender" };
        thread.Start();
        _ = Handle;
    }
    private void Run()
    {
        try
        {
            className = "EveX.Settings.Sender." + Guid.NewGuid().ToString("N");
            var windowClass = new Native.WindowClass { Size = (uint)Marshal.SizeOf<Native.WindowClass>(), Procedure = procedure, ClassName = className, Instance = Native.GetModuleHandle(null) };
            if (Native.RegisterClassEx(ref windowClass) == 0) throw new Win32Exception();
            var handle = Native.CreateWindowEx(0, className, "EVE-X-Preview Settings client", 0, 0, 0, 0, 0, 0, 0, windowClass.Instance, 0);
            if (handle == 0) throw new Win32Exception();
            ready.SetResult(handle);
            while (Native.GetMessage(out var message, 0, 0, 0) > 0) { Native.TranslateMessage(ref message); Native.DispatchMessage(ref message); }
            Native.UnregisterClass(className, windowClass.Instance);
        }
        catch (Exception e) { ready.TrySetException(e); }
    }
    private nint WindowProcedure(nint window, uint message, nint wParam, nint lParam)
    {
        if (message == 0x0010) { Native.DestroyWindow(window); return 0; }
        if (message == 0x0002) { Native.PostQuitMessage(0); return 0; }
        return Native.DefWindowProc(window, message, wParam, lParam);
    }
    public void Dispose() { if (ready.Task.IsCompletedSuccessfully) Native.PostMessage(Handle, 0x0010, 0, 0); thread.Join(1000); }
}

internal static class Native
{
    [UnmanagedFunctionPointer(CallingConvention.Winapi)] public delegate nint WindowProcedure(nint window, uint message, nint wParam, nint lParam);
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] public struct WindowClass
    {
        public uint Size, Style;
        public WindowProcedure Procedure;
        public int ClassExtra, WindowExtra;
        public nint Instance, Icon, Cursor, Background;
        public string? MenuName, ClassName;
        public nint SmallIcon;
    }
    [StructLayout(LayoutKind.Sequential)] public struct Message { public nint Window; public uint Id; public nuint WParam; public nint LParam; public uint Time; public int X, Y; public uint Private; }
    [StructLayout(LayoutKind.Sequential)] public struct CopyData { public nuint Identifier; public int ByteCount; public nint Data; }
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] public static extern ushort RegisterClassEx(ref WindowClass value);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern bool UnregisterClass(string name, nint instance);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] public static extern nint GetModuleHandle(string? name);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] public static extern nint CreateWindowEx(uint extended, string className, string title, uint style, int x, int y, int width, int height, nint parent, nint menu, nint instance, nint parameter);
    [DllImport("user32.dll")] public static extern bool DestroyWindow(nint window);
    [DllImport("user32.dll")] public static extern int GetMessage(out Message message, nint window, uint min, uint max);
    [DllImport("user32.dll")] public static extern bool TranslateMessage(ref Message message);
    [DllImport("user32.dll")] public static extern nint DispatchMessage(ref Message message);
    [DllImport("user32.dll")] public static extern void PostQuitMessage(int exitCode);
    [DllImport("user32.dll")] public static extern bool PostMessage(nint window, uint message, nint wParam, nint lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern nint DefWindowProc(nint window, uint message, nint wParam, nint lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern nint FindWindow(string? className, string title);
    [DllImport("user32.dll", SetLastError = true)] public static extern nint SendMessageTimeout(nint window, uint message, nint sender, ref CopyData data, uint flags, uint timeout, out nuint result);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(nint window, out uint processId);
    [DllImport("kernel32.dll", SetLastError = true)] public static extern nint OpenProcess(uint access, bool inherit, uint processId);
    [DllImport("advapi32.dll", SetLastError = true)] public static extern bool OpenProcessToken(nint process, uint access, out nint token);
    [DllImport("kernel32.dll")] public static extern bool CloseHandle(nint handle);
}
