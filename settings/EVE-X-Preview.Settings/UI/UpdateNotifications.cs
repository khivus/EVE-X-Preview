using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using EveX.Settings.Bridge;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    private readonly Button releaseNotice=new() {IsVisible=false,Padding=new Thickness(6,3),FontSize=12,Content=new TextBlock()};
    private readonly DispatcherTimer updatePoll=new() {Interval=TimeSpan.FromMilliseconds(500)};
    private bool pollingUpdates;

    private async Task StartBackgroundUpdatesAsync()
    {
        if(Program.CaptureDirectory is not null || !session.Connected)return;
        try {
            await session.ReadAsync("command",new(){["name"]="check-updates",["arguments"]=new JsonObject {["manual"]=false}});
            await PollUpdatesAsync();
        } catch(Exception error) {System.Diagnostics.Trace.WriteLine("Background update check: "+error.Message);}
    }
    private async Task PollUpdatesAsync()
    {
        if(closing || !session.Connected) {updatePoll.Stop();return;}
        if(pollingUpdates)return;
        pollingUpdates=true;
        try {
            var status=await session.ReadAsync("command",new(){["name"]="update-status",["arguments"]=new JsonObject()});
            SetAboutStatus(status);
            if(status.Bool("busy"))updatePoll.Start();else updatePoll.Stop();
        } catch(Exception error) {updatePoll.Stop();System.Diagnostics.Trace.WriteLine("Background update status: "+error.Message);}
        finally {pollingUpdates=false;}
    }
    private void RefreshReleaseNotice()
    {
        releaseNotice.IsVisible=aboutStatus.Bool("updateAvailable");
        releaseNotice.IsEnabled=session.Connected;
        ((TextBlock)releaseNotice.Content!).Text="Update v"+aboutStatus.Text("stable").TrimStart('v')+" available";
        ((TextBlock)releaseNotice.Content!).TextWrapping=Avalonia.Media.TextWrapping.Wrap;
        ToolTip.SetTip(releaseNotice,"Open About to review the new stable release. Settings are not installed automatically.");
    }
}
