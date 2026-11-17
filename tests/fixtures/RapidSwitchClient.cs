using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

// Test-only receiver. Observe the native message queue without AutoHotkey's
// OnMessage callback thread limit, which can skip a second key-down while
// the preceding Shift-down callback is still logging.
internal sealed class RapidSwitchClient : Form, IMessageFilter
{
    [DllImport("user32.dll")] private static extern short GetKeyState(int key);
    [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
    [DllImport("user32.dll")] private static extern IntPtr GetMessageExtraInfo();
    [DllImport("kernel32.dll")] private static extern ulong GetTickCount64();

    private readonly int clientId;
    private readonly int logDelay;
    private readonly StreamWriter log;

    private RapidSwitchClient(int id, string output, int delay)
    {
        clientId = id;
        logDelay = delay;
        log = new StreamWriter(new FileStream(Path.Combine(output, "client-" + id + ".csv"),
            FileMode.CreateNew, FileAccess.Write, FileShare.ReadWrite));
        log.AutoFlush = true;
        Text = "EVE - Rapid Pilot " + id.ToString("00");
        BackColor = Color.FromArgb(24, 40, 56);
        ForeColor = Color.White;
        StartPosition = FormStartPosition.Manual;
        Location = new Point(150, 100);
        ClientSize = new Size(850, 600);
        Application.AddMessageFilter(this);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        e.Graphics.DrawString("Input receiver " + clientId, Font, Brushes.White, 20, 20);
    }

    public bool PreFilterMessage(ref Message message)
    {
        if (message.Msg == 0x100 || message.Msg == 0x101 ||
            message.Msg == 0x104 || message.Msg == 0x105)
        {
            int key = unchecked((int)message.WParam.ToInt64());
            if (key == 0x46 || key == 0x10 || key == 0x7C)
                Record(message.Msg + ":" + key);
        }
        return false; // Record input without consuming or synthesizing it.
    }

    protected override void WndProc(ref Message message)
    {
        if (message.Msg == 0x6 && log != null)
            Record("activate:" + (message.WParam.ToInt64() & 65535));
        base.WndProc(ref message);
    }

    private void Record(string kind)
    {
        // Snapshot before any delayed IO. No nested message pump: subsequent
        // input remains in the OS queue even when this write is slow.
        string row = GetTickCount64() + "," + clientId + "," + kind + "," +
            ((GetKeyState(0x10) & 0x8000) != 0 ? 1 : 0) + "," +
            ((GetAsyncKeyState(0x10) & 0x8000) != 0 ? 1 : 0) + "," +
            GetMessageExtraInfo().ToInt64();
        if (logDelay != 0) Thread.Sleep(logDelay);
        log.WriteLine(row);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) Application.RemoveMessageFilter(this);
        base.Dispose(disposing);
        if (disposing && log != null) log.Dispose();
    }

    [STAThread]
    private static int Main(string[] args)
    {
        int id;
        if (!string.Equals(Environment.UserName, "WDAGUtilityAccount", StringComparison.OrdinalIgnoreCase) ||
            !AppDomain.CurrentDomain.BaseDirectory.StartsWith(@"C:\Checks\", StringComparison.OrdinalIgnoreCase) ||
            args.Length < 2 || !int.TryParse(args[0], out id) || id < 1 || id > 32 ||
            !Path.GetFullPath(args[1]).StartsWith(@"C:\Output\", StringComparison.OrdinalIgnoreCase))
            return 2;
        int delay = args.Length == 3 && args[2] == "--slow-log" ? 25 : 0;
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
        try
        {
            Application.EnableVisualStyles();
            using (var window = new RapidSwitchClient(id, args[1], delay))
                Application.Run(window);
            return 0;
        }
        catch (Exception error)
        {
            File.WriteAllText(Path.Combine(args[1], "receiver-error-" + id + ".txt"), error.ToString());
            return 1;
        }
    }
}
