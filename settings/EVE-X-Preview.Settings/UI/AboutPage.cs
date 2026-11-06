using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using EveX.Settings.Bridge;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    private JsonObject aboutStatus=new();
    private readonly Dictionary<string,TextBlock> versionLabels=new();
    private readonly Dictionary<string,Button> aboutButtons=new();
    private TextBlock? versionStatus;

    private void RenderAbout(int generation)
    {
        versionLabels.Clear();aboutButtons.Clear();
        var panel=new StackPanel {Spacing=12};
        panel.Children.Add(new TextBlock {Text="EVE-X-Preview",FontWeight=FontWeight.SemiBold});
        foreach(var (key,label) in new[]{("installed","Installed version"),("stable","Latest stable"),("prerelease","Latest pre-release")}) {
            var row=new Grid {ColumnDefinitions=new ColumnDefinitions("178,*")};
            row.Children.Add(SearchText(label));var value=new TextBlock();versionLabels[key]=value;
            Grid.SetColumn(value,1);row.Children.Add(value);panel.Children.Add(row);
        }
        panel.Children.Add(Muted("Created by gonzo83. Maintained by khivus.\nContributions by mrmjstc and CJKondur."));
        versionStatus=Muted("");panel.Children.Add(versionStatus);
        var actions=new StackPanel {Spacing=10,Width=380,HorizontalAlignment=HorizontalAlignment.Left};
        foreach(var id in new[]{"check-updates","install-stable","install-prerelease","download-settings","changelog","documentation","report-problem"}) {
            var command=session.Metadata!.Commands.FirstOrDefault(c=>c.Text("id")==id);if(command is null)continue;
            var button=ActionButton(command.Text("label"),()=>RunCommandAsync(command),false);
            button.HorizontalAlignment=HorizontalAlignment.Stretch;button.HorizontalContentAlignment=HorizontalAlignment.Center;
            aboutButtons[id]=button;actions.Children.Add(button);
        }
        panel.Children.Add(actions);page.Children.Add(panel);RefreshAboutControls();LoadAboutStatus(generation);
    }
    private async void LoadAboutStatus(int generation)
    {
        try {
            var result=await session.ReadAsync("command",new(){["name"]="update-status",["arguments"]=new JsonObject()});
            if(generation==pageGeneration)SetAboutStatus(result);
        } catch(Exception error) {if(generation==pageGeneration && versionStatus is not null)versionStatus.Text=error.Message;}
    }
    private void SetAboutStatus(JsonObject value) {aboutStatus=value;RefreshAboutControls();RefreshReleaseNotice();}
    private void RefreshAboutControls()
    {
        if(sectionId!="about")return;
        foreach(var (key,label) in versionLabels) {
            var tag=aboutStatus.Text(key);label.Text=tag!=""?"v"+tag.TrimStart('v'):key=="installed"?"Loading":aboutStatus.Bool("busy")?"Checking":"Not checked";
        }
        if(versionStatus is not null)versionStatus.Text=aboutStatus.Text("status","Loading version information");
        foreach(var (id,button) in aboutButtons) {
            var command=session.Metadata!.Commands.First(c=>c.Text("id")==id);
            button.IsEnabled=session.Connected&&!busy&&(!command.Bool("runtimeOnly")||session.Role!="offline");
            if(id=="download-settings") {
                button.IsEnabled &= !aboutStatus.Bool("settingsCurrent") && !aboutStatus.Bool("busy") && aboutStatus.Text("settingsUrl")!="" && !session.Dirty;
                ToolTip.SetTip(button,aboutStatus.Bool("settingsCurrent")?"The latest settings editor is already installed.":"Download beside EVE-X-Preview and open the editor.");
            }
            if(id is "install-stable" or "install-prerelease") {
                var pre=id=="install-prerelease";button.Content=aboutStatus.Text(pre?"prereleaseLabel":"stableLabel",pre?"Install pre-release (check first)":"Install stable version (check first)").TrimEnd('.');
                button.IsEnabled &= aboutStatus.Bool(pre?"prereleaseAvailable":"stableAvailable")&&!aboutStatus.Bool("busy")&&!session.Dirty;
                ToolTip.SetTip(button,session.Dirty?"Apply or discard your settings draft before installing another version.":session.Role=="offline"?"Start EVE-X-Preview to use the installer.":null);
            }
        }
    }
}
